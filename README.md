# Indian University And Course Scraper

Scrapes Indian university and course data, cleans it, and exports relational data to `university_courses.xlsx`.

## Prerequisites

- Python 3.9+
- Internet connection

## Setup Instructions

1. Open terminal in this folder:
   - `d:\pyhton\AI-ML web Scraping`
2. (Optional) Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

## Run Instructions

Run the scraper:

```bash
python university_scraper.py
```

The script prints progress logs such as:
- Which university is being scraped
- How many course links were found
- Which courses were added
- Final export summary

## Output

Generated file:
- `university_courses.xlsx`

Sheets:
- `Universities`
  - `university_id`
  - `university_name`
  - `country`
  - `city`
  - `website`
- `Courses`
  - `course_id`
  - `university_id` (foreign key to `Universities.university_id`)
  - `course_name`
  - `level`
  - `discipline`
  - `duration`
  - `fees`
  - `eligibility`

## How It Works (Flowchart)

```mermaid
flowchart TD
    A[Start Script] --> B[Load India university target list]
    B --> C[Fetch university page from TopUniversities]
    C --> D[Extract university name, city, country]
    D --> E[Fetch Wikipedia page]
    E --> F[Extract official website and fallback location data]
    F --> G{Country is India?}
    G -- No --> H[Skip university]
    G -- Yes --> I[Collect university record with university_id]
    I --> J[Extract course/program links]
    J --> K[Open each course page]
    K --> L[Extract fields: course_name, level, discipline, duration, fees, eligibility]
    L --> M[Clean text, fill missing with N/A, deduplicate]
    M --> N{At least 5 courses collected?}
    N -- No --> K
    N -- Yes --> O[Move to next university]
    H --> O
    O --> P[Build pandas DataFrames]
    P --> Q[Validate IDs and foreign-key integrity]
    Q --> R[Export to university_courses.xlsx (2 sheets)]
    R --> S[Done]
```

## Data Quality Rules Applied

- Random delay between requests
- Rotating user-agent headers
- Missing values replaced with `N/A`
- Duplicate universities/courses removed
- No null IDs
- Course rows filtered to valid `university_id` values only
