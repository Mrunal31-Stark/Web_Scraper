import json
import random
import re
import time
from urllib.parse import urljoin, urlparse

import openpyxl  # noqa: F401 - required by pandas Excel writer engine
import pandas as pd
import requests
from bs4 import BeautifulSoup


BASE_URL = "https://www.topuniversities.com"
OUTPUT_FILE = "university_courses.xlsx"
MIN_COURSES_PER_UNIVERSITY = 5
TARGET_COUNTRY = "India"

USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 "
    "(KHTML, like Gecko) Version/17.1 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
]

UNIVERSITY_TARGETS = [
    {
        "slug": "university-hyderabad",
        "wiki_url": "https://en.wikipedia.org/wiki/University_of_Hyderabad",
    },
    {
        "slug": "osmania-university",
        "wiki_url": "https://en.wikipedia.org/wiki/Osmania_University",
    },
    {
        "slug": "university-madras",
        "wiki_url": "https://en.wikipedia.org/wiki/University_of_Madras",
    },
    {
        "slug": "national-institute-technology-warangal",
        "wiki_url": "https://en.wikipedia.org/wiki/National_Institute_of_Technology,_Warangal",
    },
    {
        "slug": "national-institute-technology-calicut",
        "wiki_url": "https://en.wikipedia.org/wiki/National_Institute_of_Technology_Calicut",
    },
    {
        "slug": "indian-institute-technology-mandi",
        "wiki_url": "https://en.wikipedia.org/wiki/Indian_Institute_of_Technology_Mandi",
    },
]

ALLOWED_LEVEL_SEGMENTS = {
    "undergrad",
    "bachelors",
    "masters",
    "phd",
    "postgrad",
    "mba",
    "diploma",
    "cert",
    "foundation",
}

LEVEL_BY_SEGMENT = {
    "undergrad": "Bachelor",
    "bachelors": "Bachelor",
    "masters": "Master",
    "postgrad": "Master",
    "phd": "PhD",
    "mba": "MBA",
    "diploma": "Diploma",
    "cert": "Certificate",
    "foundation": "Foundation",
}

COUNTRY_ALIASES = {
    "Us": "United States",
    "U.S.": "United States",
    "Usa": "United States",
    "Uk": "United Kingdom",
    "England": "United Kingdom",
}


def polite_delay(low=0.8, high=1.8):
    time.sleep(random.uniform(low, high))


def clean_text(value):
    if value is None:
        return "N/A"
    text = re.sub(r"\s+", " ", str(value)).strip()
    return text if text else "N/A"


def smart_title_case(text):
    text = clean_text(text)
    if text == "N/A":
        return text

    small_words = {"and", "or", "of", "the", "in", "on", "for", "to", "at", "by", "with"}
    words = []
    for i, word in enumerate(text.split()):
        if word.isupper() and len(word) <= 5:
            words.append(word)
        elif i > 0 and word.lower() in small_words:
            words.append(word.lower())
        else:
            words.append(word.capitalize())
    titled = " ".join(words)
    return COUNTRY_ALIASES.get(titled, titled)


def build_headers():
    return {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept-Language": "en-US,en;q=0.9",
    }


def fetch_html(session, url, retries=3, timeout=30):
    for attempt in range(1, retries + 1):
        try:
            polite_delay()
            response = session.get(url, headers=build_headers(), timeout=timeout)
            if response.status_code == 200:
                return response.text
            print(f"  Request failed ({response.status_code}) for {url} [attempt {attempt}/{retries}]")
        except requests.RequestException as exc:
            print(f"  Request error for {url}: {exc} [attempt {attempt}/{retries}]")
        time.sleep(0.6 * attempt)
    return None


def parse_jsonld_objects(soup):
    objects = []
    for script_tag in soup.select('script[type="application/ld+json"]'):
        raw = script_tag.get_text(strip=True)
        if not raw:
            continue
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            continue
        if isinstance(parsed, list):
            objects.extend(parsed)
        else:
            objects.append(parsed)
    return objects


def extract_city_country_from_wikipedia(soup):
    infobox = soup.select_one("table.infobox")
    if not infobox:
        return "N/A", "N/A"

    location_raw = "N/A"
    for row in infobox.select("tr"):
        header = row.find("th")
        cell = row.find("td")
        if not header or not cell:
            continue
        label = clean_text(header.get_text(" ", strip=True)).lower()
        if label in {"location", "address"}:
            location_raw = clean_text(cell.get_text(" ", strip=True))
            break

    if location_raw == "N/A":
        return "N/A", "N/A"

    cleaned = re.sub(r"\[[^\]]*\]", "", location_raw)
    parts = [clean_text(p) for p in cleaned.split(",")]
    filtered_parts = []
    for part in parts:
        if part == "N/A":
            continue
        if re.search(r"\d", part):
            continue
        if re.search(r"\b[NSWE]\b", part):
            continue
        filtered_parts.append(part)

    if not filtered_parts:
        return "N/A", "N/A"
    if len(filtered_parts) == 1:
        return smart_title_case(filtered_parts[0]), "N/A"

    city = smart_title_case(filtered_parts[0])
    country = smart_title_case(filtered_parts[-1])
    return city, country


def extract_website_from_wikipedia(session, wiki_url):
    wiki_html = fetch_html(session, wiki_url)
    if not wiki_html:
        return "N/A", "N/A", "N/A", "N/A"

    soup = BeautifulSoup(wiki_html, "html.parser")
    heading = soup.select_one("#firstHeading")
    wiki_name = clean_text(heading.get_text(" ", strip=True)) if heading else "N/A"

    city, country = extract_city_country_from_wikipedia(soup)
    website = "N/A"
    infobox = soup.select_one("table.infobox")
    if infobox:
        for row in infobox.select("tr"):
            header = row.find("th")
            cell = row.find("td")
            if not header or not cell:
                continue
            label = clean_text(header.get_text(" ", strip=True)).lower()
            if label == "website":
                link = cell.find("a", href=True)
                if link:
                    href = clean_text(link["href"])
                    if href.startswith("//"):
                        href = "https:" + href
                    elif href.startswith("/"):
                        href = "https://en.wikipedia.org" + href
                    website = href
                else:
                    website = clean_text(cell.get_text(" ", strip=True))
                break

    return wiki_name, city, country, website


def extract_university_meta_from_topuniversities(html):
    soup = BeautifulSoup(html, "html.parser")
    jsonld_objects = parse_jsonld_objects(soup)

    university_name = "N/A"
    city = "N/A"
    country = "N/A"

    for obj in jsonld_objects:
        if not isinstance(obj, dict):
            continue

        if obj.get("@type") == "ProfilePage":
            main_entity = obj.get("mainEntity", {})
            if isinstance(main_entity, dict):
                university_name = clean_text(main_entity.get("name") or obj.get("name") or university_name)

        if obj.get("@type") == "CollegeOrUniversity":
            university_name = clean_text(obj.get("name") or university_name)
            departments = obj.get("department", [])
            if isinstance(departments, list):
                for department in departments:
                    if not isinstance(department, dict):
                        continue
                    address = department.get("address", {})
                    if not isinstance(address, dict):
                        continue
                    candidate_city = clean_text(address.get("addressLocality"))
                    candidate_country = clean_text(address.get("addressCountry"))
                    if candidate_city != "N/A":
                        city = smart_title_case(candidate_city)
                    if candidate_country != "N/A":
                        country = smart_title_case(candidate_country)
                    if city != "N/A" or country != "N/A":
                        break

    if university_name == "N/A":
        heading = soup.find("h1")
        if heading:
            university_name = clean_text(heading.get_text(" ", strip=True))

    return university_name, city, country


def extract_program_links(university_html, slug):
    soup = BeautifulSoup(university_html, "html.parser")
    links = []
    seen = set()

    for anchor in soup.select("a[href]"):
        href = clean_text(anchor.get("href"))
        if href == "N/A":
            continue
        href = href.split("?")[0].split("#")[0].strip()
        if not href.startswith(f"/universities/{slug}/"):
            continue

        parts = href.strip("/").split("/")
        if len(parts) < 4:
            continue
        level_segment = parts[2].lower()
        if level_segment not in ALLOWED_LEVEL_SEGMENTS:
            continue

        normalized_url = urljoin(BASE_URL, href)
        if normalized_url in seen:
            continue
        seen.add(normalized_url)
        links.append(normalized_url)

    return links


def extract_badges(soup):
    badges = {}
    for badge in soup.select("div.single-badge"):
        title_tag = badge.select_one("span.single-badge-title")
        value_tag = badge.select_one("div.badge-description h3")
        if not title_tag or not value_tag:
            continue

        title = clean_text(title_tag.get_text(" ", strip=True))
        value = clean_text(value_tag.get_text(" ", strip=True))
        if value != "N/A" and title != "N/A" and value.lower().endswith(title.lower()):
            value = clean_text(value[: -len(title)])

        badges[title.lower()] = value

    return badges


def extract_highlights(soup):
    highlights = {}
    for block in soup.select("div.prog-view-highli"):
        key_tag = block.find("h3")
        value_tag = block.find("p")
        if not key_tag or not value_tag:
            continue
        key = clean_text(key_tag.get_text(" ", strip=True)).lower()
        value = clean_text(value_tag.get_text(" ", strip=True))
        if key and key not in highlights and value != "N/A":
            highlights[key] = value
    return highlights


def extract_eligibility(soup):
    pairs = []
    for block in soup.select("div.univ-entry"):
        label_tag = block.select_one(".univ-entry-label")
        value_tag = block.select_one(".univ-entry-value")
        label = clean_text(label_tag.get_text(" ", strip=True)) if label_tag else "N/A"
        value = clean_text(value_tag.get_text(" ", strip=True)) if value_tag else "N/A"
        if label != "N/A" and value != "N/A":
            pairs.append(f"{label}: {value}")
        if len(pairs) >= 5:
            break

    if pairs:
        return "; ".join(pairs)

    for heading in soup.find_all(["h2", "h3", "h4"]):
        heading_text = clean_text(heading.get_text(" ", strip=True)).lower()
        if "admission" in heading_text or "eligibility" in heading_text or "entry requirement" in heading_text:
            candidate = heading.find_next(["p", "li", "div"])
            if candidate:
                snippet = clean_text(candidate.get_text(" ", strip=True))
                if snippet != "N/A":
                    return snippet[:220]
    return "N/A"


def url_level_segment(course_url):
    path_parts = urlparse(course_url).path.strip("/").split("/")
    if len(path_parts) >= 3:
        return path_parts[2].lower()
    return ""


def normalize_level(raw_level, course_url):
    value = clean_text(raw_level).lower()
    if "undergraduate" in value or "bachelor" in value:
        return "Bachelor"
    if "postgraduate" in value or "master" in value:
        return "Master"
    if "phd" in value or "doctoral" in value:
        return "PhD"
    if "diploma" in value:
        return "Diploma"
    if "certificate" in value:
        return "Certificate"
    if "mba" in value:
        return "MBA"

    return LEVEL_BY_SEGMENT.get(url_level_segment(course_url), "N/A")


def guess_discipline(course_name, course_url):
    course_name = clean_text(course_name)
    if course_name != "N/A":
        possible = re.sub(
            r"^(Bachelor|Master|BA|BSc|MSc|MA|PhD|Doctor of|Diploma in)\s*(of|in)?\s*",
            "",
            course_name,
            flags=re.IGNORECASE,
        ).strip()
        if possible and len(possible.split()) <= 6:
            return smart_title_case(possible)

    slug = clean_text(course_url.rstrip("/").split("/")[-1]).replace("-", " ")
    return smart_title_case(slug)


def scrape_course(session, course_url, university_id):
    html = fetch_html(session, course_url)
    if not html:
        return None

    soup = BeautifulSoup(html, "html.parser")
    heading = soup.find("h1")
    course_name = clean_text(heading.get_text(" ", strip=True)) if heading else "N/A"

    badges = extract_badges(soup)
    highlights = extract_highlights(soup)

    level = normalize_level(highlights.get("study level", "N/A"), course_url)
    discipline = clean_text(
        highlights.get("main subject")
        or badges.get("main subject area")
        or guess_discipline(course_name, course_url)
    )
    duration = clean_text(badges.get("programme duration") or badges.get("duration"))
    fees = clean_text(badges.get("tuition fee/year") or badges.get("tuition fee"))
    eligibility = clean_text(extract_eligibility(soup))

    return {
        "university_id": university_id,
        "course_name": course_name,
        "level": level,
        "discipline": smart_title_case(discipline),
        "duration": duration,
        "fees": fees,
        "eligibility": eligibility,
    }


def clean_universities(universities):
    df = pd.DataFrame(universities)
    if df.empty:
        return df

    for col in ["university_id", "university_name", "country", "city", "website"]:
        df[col] = df[col].astype(str).apply(clean_text)

    df["country"] = df["country"].apply(smart_title_case)
    df["city"] = df["city"].apply(smart_title_case)
    df["website"] = df["website"].apply(lambda x: x if x.startswith("http") else "N/A")

    df = df.drop_duplicates(subset=["university_name", "country", "city"], keep="first")
    return df[["university_id", "university_name", "country", "city", "website"]]


def clean_courses(courses):
    df = pd.DataFrame(courses)
    if df.empty:
        return df

    for col in ["university_id", "course_name", "level", "discipline", "duration", "fees", "eligibility"]:
        df[col] = df[col].astype(str).apply(clean_text)

    df["discipline"] = df["discipline"].apply(smart_title_case)

    df = df.drop_duplicates(subset=["university_id", "course_name", "level"], keep="first").reset_index(drop=True)
    df.insert(0, "course_id", [f"C{i + 1:04d}" for i in range(len(df))])
    return df[
        [
            "course_id",
            "university_id",
            "course_name",
            "level",
            "discipline",
            "duration",
            "fees",
            "eligibility",
        ]
    ]


def export_to_excel(universities_df, courses_df, output_path=OUTPUT_FILE):
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        universities_df.to_excel(writer, sheet_name="Universities", index=False)
        courses_df.to_excel(writer, sheet_name="Courses", index=False)


def main():
    print(f"Starting university and course scraping pipeline for {TARGET_COUNTRY}...")
    session = requests.Session()

    universities = []
    courses = []
    global_course_keys = set()

    for idx, target in enumerate(UNIVERSITY_TARGETS, start=1):
        university_id = f"U{idx:03d}"
        slug = target["slug"]
        topuni_url = f"{BASE_URL}/universities/{slug}"
        print(f"\nScraping university {idx}/{len(UNIVERSITY_TARGETS)}: {slug}")

        university_html = fetch_html(session, topuni_url)
        if not university_html:
            print(f"  Skipping {slug}: unable to fetch university page.")
            continue

        topuni_name, topuni_city, topuni_country = extract_university_meta_from_topuniversities(university_html)
        wiki_name, wiki_city, wiki_country, website = extract_website_from_wikipedia(session, target["wiki_url"])

        university_name = topuni_name if topuni_name != "N/A" else wiki_name
        city = topuni_city if topuni_city != "N/A" else wiki_city
        country = topuni_country if topuni_country != "N/A" else wiki_country
        normalized_country = smart_title_case(country)

        if normalized_country != TARGET_COUNTRY:
            print(f"  Skipping {slug}: country resolved as {normalized_country}.")
            continue

        university_record = {
            "university_id": university_id,
            "university_name": clean_text(university_name),
            "country": normalized_country,
            "city": smart_title_case(city),
            "website": clean_text(website),
        }
        universities.append(university_record)
        print(f"  University captured: {university_record['university_name']} ({university_record['country']})")

        program_links = extract_program_links(university_html, slug)
        print(f"  Program links discovered: {len(program_links)}")

        collected = 0
        for program_url in program_links:
            if collected >= MIN_COURSES_PER_UNIVERSITY:
                break

            try:
                course = scrape_course(session, program_url, university_id)
            except Exception as exc:
                print(f"  Failed to parse course page {program_url}: {exc}")
                continue

            if not course:
                continue

            dedupe_key = (
                course["university_id"].lower(),
                course["course_name"].lower(),
                course["level"].lower(),
            )
            if dedupe_key in global_course_keys:
                continue

            global_course_keys.add(dedupe_key)
            courses.append(course)
            collected += 1
            print(f"  Added course {collected}: {course['course_name']}")

        if collected < MIN_COURSES_PER_UNIVERSITY:
            print(
                f"  Warning: only {collected} courses were captured for {university_record['university_name']}."
            )

    universities_df = clean_universities(universities)
    courses_df = clean_courses(courses)

    if universities_df.empty:
        print("No university data was collected. Excel file was not created.")
        return

    if courses_df.empty:
        print("No course data was collected. Excel file was not created.")
        return

    # Keep only courses whose university_id exists in the universities sheet.
    valid_university_ids = set(universities_df["university_id"].tolist())
    courses_df = courses_df[courses_df["university_id"].isin(valid_university_ids)].reset_index(drop=True)
    courses_df["course_id"] = [f"C{i + 1:04d}" for i in range(len(courses_df))]

    export_to_excel(universities_df, courses_df, OUTPUT_FILE)

    print("\nScraping completed successfully.")
    print(f"Universities exported: {len(universities_df)}")
    print(f"Courses exported: {len(courses_df)}")
    print(f"Output file: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
