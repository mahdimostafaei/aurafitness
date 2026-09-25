import base64
import urllib.parse
import re
from pathlib import Path
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML

# Setup Paths
BASE_DIR = Path(__file__).parent
TEMPLATES_DIR = BASE_DIR / "templates"
MEDIA_DIR = BASE_DIR / "media"
OUTPUT_PDF = BASE_DIR / "local_test_blueprint.pdf"

jinja_env = Environment(loader=FileSystemLoader(str(TEMPLATES_DIR)))

def get_logos():
    logos = {"bw": None, "main": None}
    
    # 1. Load the black-and-white logo for headers and footers
    bw_path = MEDIA_DIR / "logo_bw.jpg" # Change to .jpg if necessary
    if bw_path.exists():
        with open(bw_path, "rb") as f:
            logos["bw"] = base64.b64encode(f.read()).decode("utf-8")
            
    # 2. Load the main colorful logo for the cover page
    main_path = MEDIA_DIR / "logo.jpg" # Change to .jpg if necessary
    if main_path.exists():
        with open(main_path, "rb") as f:
            logos["main"] = base64.b64encode(f.read()).decode("utf-8")
            
    return logos

def parse_workout_response(raw_text: str):
    def extract_section(tag: str, default: str = "") -> str:
        pattern = rf"\[SECTION:\s*{tag}\](.*?)(?=\[SECTION:|\Z)"
        match = re.search(pattern, raw_text, re.DOTALL | re.IGNORECASE)
        return match.group(1).strip() if match else default

    analysis = extract_section("ANALYSIS", "Analysis unavailable.")
    warmup = extract_section("WARMUP", "Dynamic warm-up recommended.")
    workouts_raw = extract_section("WORKOUTS", "")
    progression = extract_section("PROGRESSION", "Apply progressive overload.")

    days = []
    current_day = None

    for line in workouts_raw.splitlines():
        line = line.strip()
        if not line:
            continue

        if line.upper().startswith("DAY "):
            clean_title = re.sub(r"[#*]", "", line).strip()
            current_day = {"title": clean_title, "exercises": []}
            days.append(current_day)
        elif current_day and ("|" in line or "sets" in line.lower()):
            clean_line = re.sub(r"^Exercise\s*\d*:\s*", "", line, flags=re.I).strip()

            cue = ""
            cue_match = re.search(r"(?:Cue|Tips?)\s*:?\s*([^|]*)", clean_line, re.I)
            if cue_match:
                cue = cue_match.group(1).strip()
                clean_line = clean_line.replace(cue_match.group(0), "")

            rest = ""
            rest_match = re.search(r"Rest\s*:?\s*([^|]*)", clean_line, re.I)
            if rest_match:
                rest = rest_match.group(1).strip()
                clean_line = clean_line.replace(rest_match.group(0), "")

            parts = [p.strip() for p in clean_line.split("|") if p.strip()]
            ex_name = parts[0] if parts else "Exercise"
            reps_raw = parts[1] if len(parts) > 1 else ""

            current_day["exercises"].append({
                "name": ex_name,
                "sets_reps": reps_raw,
                "rest": rest,
                "tip": cue,
                "video_query": urllib.parse.quote(f"{ex_name} exercise form demo")
            })

    return {
        "analysis": analysis,
        "warmup": warmup,
        "days": days,
        "progression": progression
    }

def run_local_test():
    # Mock data mimicking Google Sheets input
    mock_intake = {
        "user_name": "Ali",
        "preferred_days": "3",
        "primary_goal": "Strength",
        "preferred_split": "Full Body"
    }

    # Mock AI text adhering strictly to the Master Prompt format requirements
    mock_ai_text = """
[SECTION: ANALYSIS]
Whether my chosen training frequency makes sense: 3 days per week is an excellent frequency for a beginner to intermediate trainee focusing on strength[cite: 13].
Whether my selected split is appropriate: A full-body split 3 days a week allows for high frequency of the major movement patterns, which is optimal for strength adaptations[cite: 13].
Main programming priorities: The focus is on stable, compound movements to safely build a base of strength without overcomplicating the routine[cite: 13].

[SECTION: WARMUP]
Spend 5-8 minutes performing dynamic stretches (arm circles, leg swings, and bodyweight squats) to elevate your heart rate. For the first major exercise of each workout, perform 2 lighter warm-up sets before using your working weight[cite: 13].

[SECTION: WORKOUTS]
DAY 1 — Full Body A
Goblet Squat | 3 sets × 8–12 reps | Rest: 2-3 min | Cue: Keep your chest tall and push your knees out in line with your toes.
Dumbbell Bench Press | 3 sets × 8–12 reps | Rest: 2-3 min | Cue: Keep your shoulder blades stable against the bench and control the descent.
Seated Cable Row | 3 sets × 10–15 reps | Rest: 90 sec | Cue: Pull with your elbows and squeeze your back at the top.
Bodyweight Plank | 3 sets × 30–45 sec | Rest: 60 sec | Cue: Brace your core like you are about to be punched.

DAY 2 — Full Body B
Romanian Deadlift | 3 sets × 8–12 reps | Rest: 2-3 min | Cue: Push your hips back until you feel a stretch in your hamstrings.
Overhead Dumbbell Press | 3 sets × 8–12 reps | Rest: 2-3 min | Cue: Brace your core and press straight up without leaning back.
Lat Pulldown | 3 sets × 10–15 reps | Rest: 90 sec | Cue: Pull the bar down to your upper chest.
Pallof Press | 3 sets × 10–12 reps | Rest: 60 sec | Cue: Resist rotation as you press the cable forward.

[SECTION: PROGRESSION]
Monday — Day 1
Tuesday — Rest
Wednesday — Day 2
Thursday — Rest
Friday — Day 1
Saturday — Rest
Sunday — Rest

Start with a weight that allows good technique[cite: 13]. Keep approximately 1–3 repetitions in reserve on most working sets[cite: 13]. Once you can consistently reach the top of the prescribed rep range with good technique, increase the weight gradually[cite: 13]. Do not recommend increasing weight if technique breaks down[cite: 13].
    """

    print("Parsing AI Text...")
    parsed_sections = parse_workout_response(mock_ai_text)
    
    print("Loading Logos...")
    logos = get_logos()

    print("Rendering HTML Template...")
    template = jinja_env.get_template("pdf_template.html")
    rendered_html = template.render(
        intake=mock_intake,
        sections=parsed_sections,
        logo_bw=logos["bw"],       # Passes the BW logo
        logo_main=logos["main"]    # Passes the Main logo
    )

    print(f"Drawing PDF to {OUTPUT_PDF}...")
    HTML(string=rendered_html, base_url=str(TEMPLATES_DIR)).write_pdf(OUTPUT_PDF)
    print("Success! Open local_test_blueprint.pdf to view the result.")

if __name__ == "__main__":
    run_local_test()