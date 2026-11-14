#!/usr/bin/env python3
"""Scraper des agences Orpi -> xlsx (feuille Agences + feuille Equipes).

Découverte : pages /agences-immobilieres-*/ (régions, départements, villes)
             + sitemap.xml si disponible.
Détail     : page de chaque agence (nom, services, adresse, tél, équipe).
Reprise    : cache JSONL (data/cache_agences.jsonl) -> on peut relancer sans tout refaire.
"""
import json, os, re, sys, time, random, xml.etree.ElementTree as ET
from pathlib import Path
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor, as_completed

from bs4 import BeautifulSoup
from curl_cffi import requests as cffi
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

BASE = "https://www.orpi.com"
OUT_DIR = Path(os.getenv("OUT_DIR", "data"))
OUT_DIR.mkdir(parents=True, exist_ok=True)
CACHE = OUT_DIR / "cache_agences.jsonl"
XLSX = OUT_DIR / "agences_orpi.xlsx"
WORKERS = int(os.getenv("WORKERS", "4"))
DELAY = float(os.getenv("DELAY", "0.6"))
LIMIT = int(os.getenv("LIMIT", "0"))  # 0 = tout ; utile pour tester

# slugs du site qui ne sont PAS des agences
RESERVED = {
    "", "acheter", "louer", "vendre", "faire-gerer", "syndic", "investir", "actualites",
    "agences-immobilieres", "recherche", "mon-compte", "mes-favoris", "estimation-immobiliere",
    "estimation-immobiliere-en-ligne", "rejoignez-nous", "annonces-immobilieres", "conseils",
    "mentions-legales", "contact", "build", "static", "images", "blog", "devenir-agent",
    "immobilier-professionnel", "gestion-locative", "credit-immobilier", "assurance",
    "sitemap.xml", "robots.txt",
}
GERANT_RE = re.compile(r"directeur|directrice|dirigeant|g[ée]rant|responsable d.agence|"
                       r"pr[ée]sident|titulaire|chef d.agence", re.I)

_local = {}


def session():
    import threading
    t = threading.get_ident()
    if t not in _local:
        _local[t] = cffi.Session(impersonate="chrome")
    return _local[t]


def get(url, tries=5):
    last = None
    for i in range(tries):
        try:
            time.sleep(DELAY + random.random() * 0.4)
            r = session().get(url, timeout=40, headers={"Accept-Language": "fr-FR,fr;q=0.9"})
            if r.status_code == 200:
                return r.text
            if r.status_code == 404:
                return None
            last = f"HTTP {r.status_code}"
            if r.status_code in (403, 429, 503):
                time.sleep(5 * (i + 1) + random.random() * 5)
        except Exception as e:  # réseau, timeout...
            last = repr(e)
            time.sleep(2 * (i + 1))
    print(f"[WARN] échec {url}: {last}", file=sys.stderr)
    return None


# ----------------------------------------------------------------- découverte
def agency_slug(href):
    """/dmi/ ou /dmi/nous-contacter/ -> 'dmi' ; sinon None"""
    p = urlparse(urljoin(BASE, href))
    if p.netloc not in ("www.orpi.com", "orpi.com"):
        return None
    parts = [x for x in p.path.split("/") if x]
    if not parts or len(parts) > 2:
        return None
    if len(parts) == 2 and parts[1] != "nous-contacter":
        return None
    s = parts[0]
    if s in RESERVED or s.startswith("agences-immobilieres") or "." in s and s.endswith((".xml", ".txt", ".png")):
        return None
    return s


def discover():
    slugs, seen, queue = set(), set(), [f"{BASE}/agences-immobilieres/"]

    # 1) sitemap (rapide si dispo)
    for sm in (f"{BASE}/sitemap.xml",):
        xml = get(sm, tries=2)
        if xml:
            locs = re.findall(r"<loc>([^<]+)</loc>", xml)
            for loc in locs:
                if loc.endswith(".xml"):
                    sub = get(loc, tries=2)
                    for l2 in re.findall(r"<loc>([^<]+)</loc>", sub or ""):
                        s = agency_slug(l2)
                        if s:
                            slugs.add(s)
                else:
                    s = agency_slug(loc)
                    if s:
                        slugs.add(s)
            print(f"[sitemap] {len(slugs)} agences")

    # 2) crawl des pages de listing
    while queue:
        url = queue.pop()
        if url in seen:
            continue
        seen.add(url)
        html = get(url)
        if not html:
            continue
        soup = BeautifulSoup(html, "lxml")
        for a in soup.find_all("a", href=True):
            href = a["href"].split("#")[0]
            if not href:
                continue
            full = urljoin(BASE, href)
            path = urlparse(full).path
            if re.match(r"^/agences-immobilieres-[^/]+/?$", path) or \
               re.match(r"^/agences-immobilieres/?$", path):
                if full not in seen:
                    queue.append(full)
                continue
            s = agency_slug(full)
            if s:
                slugs.add(s)
        # pagination éventuelle
        for a in soup.select('a[rel="next"], link[rel="next"]'):
            nxt = a.get("href")
            if nxt:
                queue.append(urljoin(url, nxt))
        if len(seen) % 25 == 0:
            print(f"[crawl] {len(seen)} pages listing, {len(slugs)} agences")
    print(f"[découverte] {len(slugs)} agences, {len(seen)} pages listing")
    return sorted(slugs)


# ------------------------------------------------------------------ parsing
def clean(t):
    return re.sub(r"\s+", " ", (t or "").replace("\xa0", " ")).strip()


def parse_agency(slug, html):
    soup = BeautifulSoup(html, "lxml")
    card = soup.select_one("article.c-agency-card") or soup
    name_el = card.select_one('[itemprop="name"]') or soup.select_one('[itemprop="name"]')
    if not name_el:
        return None  # pas une page agence
    name = clean(name_el.get_text())

    services = []
    for li in card.select("ul li"):
        s = clean(li.get_text()).lstrip("-").strip()
        if re.match(r"(lundi|mardi|mercredi|jeudi|vendredi|samedi|dimanche)\b", s, re.I):
            continue  # horaires d'ouverture, pas des services
        if s and s not in services:
            services.append(s)

    def itp(prop):
        el = soup.select_one(f'[itemprop="{prop}"]')
        return clean(el.get_text()) if el else ""

    team, seen_m = [], set()
    for art in soup.select("article.c-box--senary"):
        h = art.find("h4")
        if not h:
            continue
        nom = clean(h.get_text())
        p = art.find("p")
        role = clean(p.get_text()) if p else ""
        tel_a = art.select_one('a[href^="tel:"]')
        tel = clean(tel_a["href"][4:]) if tel_a else ""
        if (nom, role) in seen_m:
            continue
        seen_m.add((nom, role))
        team.append({"nom": nom, "role": role, "telephone": tel})

    gerants = [m["nom"] for m in team if GERANT_RE.search(m["role"])]
    ref = re.search(r'&quot;agenceRef&quot;:&quot;(\w+)&quot;|"agenceRef":"(\w+)"', html)
    dept = re.search(r'departementNom\W+([^"&\\]+)', html)
    return {
        "slug": slug,
        "url": f"{BASE}/{slug}/",
        "nom_agence": name,
        "gerant": " / ".join(gerants),
        "services": " - ".join(services),
        "adresse": itp("streetAddress"),
        "code_postal": itp("postalCode"),
        "ville": itp("addressLocality"),
        "telephone": itp("telephone"),
        "ref_agence": (ref.group(1) or ref.group(2)) if ref else "",
        "equipe": team,
    }


def fetch_agency(slug):
    html = get(f"{BASE}/{slug}/")
    if not html:
        return slug, None
    return slug, parse_agency(slug, html)


# --------------------------------------------------------------------- xlsx
def write_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Agences"
    ws2 = wb.create_sheet("Equipes")
    h1 = ["Nom de l'agence", "Nom du gérant", "Services", "Adresse", "Code postal", "Ville",
          "Téléphone", "URL"]
    h2 = ["Agence", "URL agence", "Nom", "Fonction", "Téléphone"]
    ws.append(h1)
    ws2.append(h2)
    for r in rows:
        ws.append([r["nom_agence"], r["gerant"], r["services"], r["adresse"], r["code_postal"],
                   r["ville"], r["telephone"], r["url"]])
        for m in r["equipe"]:
            ws2.append([r["nom_agence"], r["url"], m["nom"], m["role"], m["telephone"]])
    fill = PatternFill("solid", fgColor="1F3864")
    for sheet, widths in ((ws, [38, 30, 60, 36, 12, 24, 16, 38]), (ws2, [38, 38, 28, 36, 16])):
        for c in sheet[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = fill
            c.alignment = Alignment(vertical="center")
        for i, w in enumerate(widths, 1):
            sheet.column_dimensions[get_column_letter(i)].width = w
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
    wb.save(XLSX)
    print(f"[xlsx] {XLSX} — {len(rows)} agences, {ws2.max_row - 1} membres")


# --------------------------------------------------------------------- main
def load_cache():
    done = {}
    if CACHE.exists():
        for line in CACHE.read_text(encoding="utf-8").splitlines():
            try:
                d = json.loads(line)
                done[d["slug"]] = d
            except Exception:
                pass
    return done


def main():
    done = load_cache()
    slugs = discover()
    if LIMIT:
        slugs = slugs[:LIMIT]
    todo = [s for s in slugs if s not in done]
    print(f"[détail] {len(todo)} à traiter ({len(done)} en cache)")
    not_agency = set()
    with ThreadPoolExecutor(max_workers=WORKERS) as ex, CACHE.open("a", encoding="utf-8") as f:
        futs = [ex.submit(fetch_agency, s) for s in todo]
        for n, fu in enumerate(as_completed(futs), 1):
            slug, data = fu.result()
            if data:
                done[slug] = data
                f.write(json.dumps(data, ensure_ascii=False) + "\n")
                f.flush()
            else:
                not_agency.add(slug)
            if n % 50 == 0:
                print(f"[détail] {n}/{len(todo)}")
    if not_agency:
        print(f"[info] {len(not_agency)} slugs ignorés (non-agence / erreur)")
    rows = sorted((done[s] for s in slugs if s in done), key=lambda r: (r["code_postal"], r["nom_agence"]))
    write_xlsx(rows)


if __name__ == "__main__":
    main()
