"""Collect entry/junior full-time tech jobs in Germany from LinkedIn (guest API),
Bundesagentur fuer Arbeit and Arbeitnow. Writes raw.json."""
import json, re, sys, time, random, html
from concurrent.futures import ThreadPoolExecutor
import requests
from bs4 import BeautifulSoup

OUT = sys.argv[1]
UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/128.0 Safari/537.36",
      "Accept-Language": "en-US,en;q=0.9"}
S = requests.Session(); S.headers.update(UA)

LI_QUERIES = [
    "Junior Software Engineer", "Junior Software Developer", "Junior Full Stack Developer",
    "Junior Backend Developer", "Junior Frontend Developer", "Junior Python Developer",
    "Junior AI Engineer", "Junior Machine Learning Engineer", "AI Engineer", "LLM Engineer",
    "Generative AI Engineer", "Graduate Software Engineer", "Entry Level Software Engineer",
    "Junior Data Engineer", "Junior DevOps Engineer", "Junior Cloud Engineer",
    "Junior Node.js Developer", "Junior TypeScript Developer", "Junior React Developer",
    "Junior Angular Developer", "Blockchain Developer", "Junior iOS Developer",
    "Flutter Developer", "Junior Softwareentwickler", "Softwareentwickler Berufseinsteiger",
    "Trainee Software Development", "Full Stack Engineer", "Software Engineer", "Python Developer",
    "Machine Learning Engineer", "Data Scientist Junior", "Junior Platform Engineer",
    "Junior Application Security Engineer", "Junior Integration Engineer", "Graduate Program IT",
]


def li_search(q, start):
    url = "https://www.linkedin.com/jobs-guest/jobs/api/seeMoreJobPostings/search"
    p = {"keywords": q, "location": "Germany", "geoId": "101282230", "f_E": "2,3", "f_JT": "F",
         "f_TPR": "r2592000", "start": start}
    for attempt in range(6):
        try:
            r = S.get(url, params=p, timeout=30)
        except Exception:
            time.sleep(5 * (attempt + 1)); continue
        if r.status_code == 200:
            return r.text
        if r.status_code == 400:
            return ""
        time.sleep(3 * (attempt + 1) + random.random() * 3)
    return ""


def li_parse_cards(text):
    s = BeautifulSoup(text, "html.parser")
    out = []
    for c in s.select("div.base-card"):
        urn = c.get("data-entity-urn", "")
        jid = urn.split(":")[-1]
        if not jid:
            continue
        t = lambda sel: (c.select_one(sel).get_text(" ", strip=True) if c.select_one(sel) else "")
        tm = c.select_one("time")
        a = c.select_one("a.base-card__full-link")
        sal = t(".job-search-card__salary-info")
        out.append({"source": "LinkedIn", "id": "li-" + jid, "title": t("h3.base-search-card__title"),
                    "company": t("h4.base-search-card__subtitle"), "location": t(".job-search-card__location"),
                    "posted": tm.get("datetime", "") if tm else "", "salary": sal,
                    "url": f"https://www.linkedin.com/jobs/view/{jid}/",
                    "company_url": (c.select_one("h4 a") or {}).get("href", "").split("?")[0] if c.select_one("h4 a") else ""})
    return out


def li_detail(job):
    jid = job["id"][3:]
    for attempt in range(5):
        try:
            r = S.get(f"https://www.linkedin.com/jobs-guest/jobs/api/jobPosting/{jid}", timeout=30)
        except Exception:
            time.sleep(3); continue
        if r.status_code == 200:
            s = BeautifulSoup(r.text, "html.parser")
            for li in s.select("li.description__job-criteria-item"):
                k = li.select_one("h3").get_text(strip=True) if li.select_one("h3") else ""
                v = li.select_one("span").get_text(" ", strip=True) if li.select_one("span") else ""
                job[{"Seniority level": "seniority", "Employment type": "employment_type",
                     "Job function": "job_function", "Industries": "industry"}.get(k, k)] = v
            d = s.select_one("div.show-more-less-html__markup")
            job["description"] = d.get_text("\n", strip=True) if d else ""
            ap = s.select_one(".num-applicants__caption")
            job["applicants"] = ap.get_text(strip=True) if ap else ""
            return job
        if r.status_code in (404, 410):
            return job
        time.sleep(4 * (attempt + 1) + random.random() * 4)
    return job


CK = OUT + ".ck.json"
import os
def load_ck():
    return json.load(open(CK)) if os.path.exists(CK) else {"an": None, "li": {}, "done_q": [], "details": {}}
def save_ck(ck):
    json.dump(ck, open(CK + ".tmp", "w"), ensure_ascii=False); os.replace(CK + ".tmp", CK)

def linkedin(ck):
    jobs = ck["li"]
    for q in LI_QUERIES:
        if q in ck["done_q"]:
            continue
        empty = 0
        for start in range(0, 400, 10):
            cards = li_parse_cards(li_search(q, start))
            new = [c for c in cards if c["id"] not in jobs]
            for c in cards:
                c["query"] = q
                jobs.setdefault(c["id"], c)
            if not cards:
                break
            empty = empty + 1 if not new else 0
            if empty >= 3:
                break
            time.sleep(2.5 + random.random() * 2)
        ck["done_q"].append(q); save_ck(ck)
        print(f"[LI] {q}: total {len(jobs)}", flush=True)
    return list(jobs.values())


BA_QUERIES = ["Junior Softwareentwickler", "Junior Software Engineer", "Junior Developer",
              "Softwareentwickler Berufseinsteiger", "Junior Full Stack", "Junior Python",
              "Junior Data Engineer", "Junior DevOps", "AI Engineer", "Machine Learning Engineer",
              "Junior Frontend", "Junior Backend", "Absolvent Informatik", "Trainee IT",
              "Junior Webentwickler", "Junior Cloud Engineer", "Junior Anwendungsentwickler"]


def ba():
    h = {"X-API-Key": "jobboerse-jobsuche",
         "User-Agent": "Jobsuche/2.9.2 (de.arbeitsagentur.jobboerse; build:1077; iOS 15.1.0) Alamofire/5.4.4"}
    jobs = {}
    for q in BA_QUERIES:
        for page in range(1, 6):
            r = requests.get("https://rest.arbeitsagentur.de/jobboerse/jobsuche-service/pc/v4/app/jobs",
                             headers=h, timeout=30,
                             params={"was": q, "angebotsart": 1, "arbeitszeit": "vz", "veroeffentlichtseit": 30,
                                     "size": 100, "page": page})
            if r.status_code != 200:
                break
            lst = r.json().get("stellenangebote", []) or r.json().get("ergebnisliste", [])
            for j in lst:
                ref = j.get("refnr")
                if not ref or ref in jobs:
                    continue
                o = j.get("arbeitsort", {}) or {}
                ext = j.get("externeUrl") or ""
                jobs[ref] = {"source": "Bundesagentur für Arbeit", "id": "ba-" + ref, "title": j.get("titel", ""),
                             "company": j.get("arbeitgeber", ""),
                             "location": ", ".join(x for x in [o.get("ort", ""), o.get("region", "")] if x),
                             "posted": j.get("aktuelleVeroeffentlichungsdatum", "") or j.get("eintrittsdatum", ""),
                             "url": f"https://www.arbeitsagentur.de/jobsuche/jobdetail/{ref}", "apply_url": ext,
                             "employment_type": "Full-time", "query": q, "description": j.get("beruf", "")}
            if len(lst) < 100:
                break
            time.sleep(0.5)
        print(f"[BA] {q}: total {len(jobs)}", flush=True)
    return list(jobs.values())


def arbeitnow():
    jobs = []
    for page in range(1, 40):
        r = requests.get("https://www.arbeitnow.com/api/job-board-api", params={"page": page}, timeout=30)
        if r.status_code != 200:
            break
        data = r.json().get("data", [])
        if not data:
            break
        for j in data:
            d = BeautifulSoup(j.get("description", ""), "html.parser").get_text("\n", strip=True)
            jobs.append({"source": "Arbeitnow", "id": "an-" + j["slug"], "title": j["title"],
                         "company": j["company_name"], "location": j.get("location", ""),
                         "posted": time.strftime("%Y-%m-%d", time.gmtime(j.get("created_at", 0))),
                         "url": j["url"], "remote_flag": j.get("remote"), "tags": ", ".join(j.get("tags", [])),
                         "employment_type": ", ".join(j.get("job_types", [])), "description": d})
        time.sleep(0.4)
    print(f"[AN] total {len(jobs)}", flush=True)
    return jobs


if __name__ == "__main__":
    ck = load_ck()
    if ck["an"] is None:
        ck["an"] = arbeitnow(); save_ck(ck)
    allj = list(ck["an"])
    li = linkedin(ck)
    todo = [j for j in li if j["id"] not in ck["details"]]
    print(f"fetching {len(todo)} LinkedIn details", flush=True)
    with ThreadPoolExecutor(3) as ex:
        for i, j in enumerate(ex.map(li_detail, todo)):
            ck["details"][j["id"]] = j
            if i % 50 == 0:
                save_ck(ck); print(f"  detail {i}", flush=True)
    save_ck(ck)
    allj += [ck["details"].get(j["id"], j) for j in li]
    json.dump(allj, open(OUT, "w"), ensure_ascii=False)
    print("done", len(allj))
