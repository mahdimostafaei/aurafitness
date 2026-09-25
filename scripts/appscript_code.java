var FOLDER_ID_ARCHIVE = "1gl94lHos4VmEjU1jEqP1tW4JTgvBCHGF";

function doPost(e) {
  var output_headers = {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type"
  };

  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    var row_data = [
      new Date(), 
      data.user_name || "",
      data.user_age || "",
      data.user_sex || "",
      data.user_height || "",
      data.user_weight || "",
      data.user_occupation || "",
      data.daily_activity || "",
      data.daily_steps || "",
      data.training_duration || "",
      data.experience_level || "",
      data.consistent_training || "",
      data.training_hiatus || "",
      data.strength_level || "",
      data.primary_goal || "",
      data.secondary_goal || "",
      data.top_priority || "",
      data.preferred_days || "",
      data.max_days || "",
      data.workout_duration || "",
      data.preferred_split || "",
      data.training_location || "",
      data.equipment_available || "",
      data.exercises_enjoyed || "",
      data.exercises_disliked || "",
      data.exercises_cannot_perform || "",
      data.exercises_wanted || "",
      data.injuries_limitations || "",
      data.preferred_style || "",
      data.workout_difficulty || "",
      data.exercise_variety || "",
      "",            // Column AF: generated_workout
      "",            // Column AG: pdf_url
      "PENDING"      // Column AH: processing_status
    ];
    
    sheet.appendRow(row_data);

    return ContentService.createTextOutput(JSON.stringify({
      "status": "success", 
      "message": "Submission received! Your custom protocol is compiling and will be ready shortly."
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      "status": "error", 
      "message": error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function process_workout_queue() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data_range = sheet.getDataRange();
  var rows = data_range.getValues();
  
  if (rows.length <= 1) return;

  var api_key = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!api_key) throw new Error("GEMINI_API_KEY property is missing.");

  for (var i = 1; i < rows.length; i++) {
    var status = rows[i][33]; // Column AH (index 33)
    
    if (status === "PENDING") {
      var row_index = i + 1;
      sheet.getRange(row_index, 34).setValue("PROCESSING");
      SpreadsheetApp.flush();

      try {
        var user_payload = {
          user_name: rows[i][1],
          user_age: rows[i][2],
          user_sex: rows[i][3],
          user_height: rows[i][4],
          user_weight: rows[i][5],
          user_occupation: rows[i][6],
          daily_activity: rows[i][7],
          daily_steps: rows[i][8],
          training_duration: rows[i][9],
          experience_level: rows[i][10],
          consistent_training: rows[i][11],
          training_hiatus: rows[i][12],
          strength_level: rows[i][13],
          primary_goal: rows[i][14],
          secondary_goal: rows[i][15],
          top_priority: rows[i][16],
          preferred_days: rows[i][17],
          max_days: rows[i][18],
          workout_duration: rows[i][19],
          preferred_split: rows[i][20],
          training_location: rows[i][21],
          equipment_available: rows[i][22],
          exercises_enjoyed: rows[i][23],
          exercises_disliked: rows[i][24],
          exercises_cannot_perform: rows[i][25],
          exercises_wanted: rows[i][26],
          injuries_limitations: rows[i][27],
          preferred_style: rows[i][28],
          workout_difficulty: rows[i][29],
          exercise_variety: rows[i][30]
        };

        var generated_text = request_gemini_workout(user_payload, api_key);
        sheet.getRange(row_index, 32).setValue(generated_text);

        var pdf_file = build_html_pdf(user_payload, generated_text);
        var pdf_url = pdf_file.getUrl();
        sheet.getRange(row_index, 33).setValue(pdf_url);

        sheet.getRange(row_index, 34).setValue("COMPLETED");

      } catch (err) {
        sheet.getRange(row_index, 34).setValue("ERROR: " + err.toString());
      }
      SpreadsheetApp.flush();
    }
  }
}

function request_gemini_workout(data, api_key) {
  // Set to 'false' only when you have linked a billing account to lift the 20-request limit
  var TEST_MODE = true; 

  if (TEST_MODE) {
    Utilities.sleep(2000); // Simulate API latency
    return `[SECTION: ANALYSIS]
Based on your goals and schedule, a 2-day full-body split is highly effective. We will prioritize compound movements to maximize efficiency and stimulate muscle growth while managing fatigue. Because you are a beginner, we will focus on stable, easy-to-learn movements without unnecessary isolation exercises.

[SECTION: WARMUP]
Spend 5-8 minutes performing dynamic stretches (arm circles, leg swings, and bodyweight squats) to elevate your heart rate and prepare your joints. Perform 1-2 light sets of your first major exercise before using your working weight.

[SECTION: WORKOUTS]
DAY 1 — Full Body A
Exercise 1: Goblet Squat | 3 sets × 8–12 reps | Rest: 2 min | Cue: Keep your chest tall and push your knees out in line with your toes.
Exercise 2: Dumbbell Bench Press | 3 sets × 8–12 reps | Rest: 2 min | Cue: Keep your shoulder blades stable against the bench and control the descent.
Exercise 3: Seated Cable Row | 3 sets × 10–15 reps | Rest: 90 sec | Cue: Pull with your elbows and squeeze your back at the top.
Exercise 4: Bodyweight Plank | 3 sets × 30–45 sec | Rest: 60 sec | Cue: Brace your core like you are about to be punched.

DAY 2 — Full Body B
Exercise 1: Romanian Deadlift | 3 sets × 8–12 reps | Rest: 2 min | Cue: Push your hips back until you feel a stretch in your hamstrings.
Exercise 2: Overhead Dumbbell Press | 3 sets × 8–12 reps | Rest: 2 min | Cue: Brace your core and press straight up without leaning back.
Exercise 3: Lat Pulldown | 3 sets × 10–15 reps | Rest: 90 sec | Cue: Pull the bar down to your upper chest.
Exercise 4: Pallof Press | 3 sets × 10–12 reps | Rest: 60 sec | Cue: Resist rotation as you press the band/cable forward.

[SECTION: PROGRESSION]
**Weekly Structure:**
Monday: Day 1 | Tuesday: Rest | Wednesday: Rest | Thursday: Day 2 | Friday, Saturday, Sunday: Rest

**Progression Rule:**
Start with a weight that allows good technique. Keep approximately 1–3 repetitions in reserve on most working sets. Once you can consistently reach the top of the prescribed rep range with good technique, increase the weight gradually. Do not recommend increasing weight if technique breaks down.`;
  }

  var endpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-3.6-flash:generateContent?key=" + api_key;
  
  var master_prompt = `You are an experienced strength-training and fitness-programming assistant.

Your job is to create a practical, beginner-friendly gym workout program based on the person's information below.

IMPORTANT:
- This is a general fitness plan, not medical advice.
- Do not diagnose injuries, illnesses, or medical conditions.
- If the person reports an injury, significant pain, medical condition, or other issue that could affect exercise selection or safety, clearly flag it and recommend consulting an appropriate qualified professional before training around it.
- Do not recommend unnecessarily advanced or technically complicated exercises to beginners.
- Prioritize safe technique, progressive overload, recoverability, and consistency.
- The goal is to create a realistic program that the person can actually follow.

==================================================
PERSONAL INFORMATION
==================================================
Name: ${data.user_name || "Not provided"}
Age: ${data.user_age || "Not provided"}
Sex: ${data.user_sex || "Not provided"}
Height: ${data.user_height || "Not provided"} cm
Weight: ${data.user_weight || "Not provided"} kg
Occupation: ${data.user_occupation || "Not provided"}
Typical daily activity: ${data.daily_activity || "Not provided"}
Average daily steps, if known: ${data.daily_steps || "Not provided"}

==================================================
TRAINING EXPERIENCE
==================================================
How long have you been training? ${data.training_duration || "Not provided"}
Current training experience: ${data.experience_level || "Not provided"}
Have you trained consistently before? ${data.consistent_training || "Not provided"}
Have you stopped training for a long period before? ${data.training_hiatus || "Not provided"}
Current strength level, if known: ${data.strength_level || "Not provided"}

==================================================
GOAL
==================================================
Primary goal: ${data.primary_goal || "Not provided"}
Secondary goal: ${data.secondary_goal || "Not provided"}
What is most important to me? ${data.top_priority || "Not provided"}

==================================================
TRAINING SCHEDULE
==================================================
Preferred number of gym days per week: ${data.preferred_days || "Not provided"}
Maximum realistic number of gym days per week: ${data.max_days || "Not provided"}
Approximate workout duration: ${data.workout_duration || "Not provided"}
Preferred training split: ${data.preferred_split || "Not provided"}

IMPORTANT:
If I selected a split but it is not appropriate for my experience level, schedule, recovery ability, or goals, explain why and recommend a better structure. If I selected "I don't know", choose the most appropriate split for me. Do not assume that a higher-number split is automatically better.

==================================================
GYM / EQUIPMENT
==================================================
Where do I train? ${data.training_location || "Not provided"}
Equipment available: ${data.equipment_available || "Not provided"}

==================================================
EXERCISE PREFERENCES
==================================================
Exercises I enjoy: ${data.exercises_enjoyed || "Not provided"}
Exercises I dislike: ${data.exercises_disliked || "Not provided"}
Exercises I cannot perform: ${data.exercises_cannot_perform || "Not provided"}
Exercises I specifically want included: ${data.exercises_wanted || "Not provided"}

==================================================
INJURIES / LIMITATIONS / PAIN
==================================================
Details: ${data.injuries_limitations || "None reported"}

==================================================
TRAINING STYLE
==================================================
Preferred workout style: ${data.preferred_style || "Not provided"}
How challenging do I want the workouts to feel? ${data.workout_difficulty || "Not provided"}
Do I want: ${data.exercise_variety || "Not provided"}

==================================================
OUTPUT REQUIREMENTS
==================================================
Create a complete workout program based on the information above. You must wrap your responses in the exact section tags below so our system can parse them.

[SECTION: ANALYSIS]
Briefly analyze my situation and explain:
1. Whether my chosen training frequency makes sense.
2. Whether my selected split is appropriate.
3. What the main programming priorities should be for someone with my experience and goals.
4. Any important limitations or considerations.

[SECTION: WARMUP]
Provide a short general warm-up recommendation. Do not make the warm-up unnecessarily long. For the first major exercise of each workout, explain whether additional lighter warm-up sets would be useful.

[SECTION: WORKOUTS]
The workout should be organized by DAY.
For every day, use this format EXACTLY, using the pipe symbol (|) to separate metadata:

DAY 1 — [Muscle Groups / Workout Name]
Exercise 1: [Exercise name] | [Sets] sets × [Rep range] reps | Rest: [Time] | Cue: [Short technique cue]
Exercise 2: [Exercise name] | [Sets] sets × [Rep range] reps | Rest: [Time] | Cue: [Short technique cue]

Balance the program across major muscle groups without forcing every muscle into every workout. Choose sets and rep ranges based on the exercise and my goal.

[SECTION: PROGRESSION]
Show how the workouts should be repeated during the week.
Include a simple progression system (e.g., Start with a weight that allows good technique. Keep 1-3 reps in reserve. Increase weight gradually once top of rep range is hit). Provide final safety/technique notes. Do not include nutrition or a diet plan unless I explicitly ask for it.`;

  var payload = {
    "contents": [{ "parts": [{ "text": master_prompt }] }],
    "generationConfig": { "maxOutputTokens": 4096 }
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  var responseCode = 0;
  var responseText = "";
  var success = false;

  // Heavy-duty backoff loop: 5 attempts (waits 4s, 8s, 16s, 32s)
  for (var attempt = 1; attempt <= 5; attempt++) {
    var response = UrlFetchApp.fetch(endpoint, options);
    responseCode = response.getResponseCode();
    responseText = response.getContentText();
    
    if (responseCode === 200) {
      success = true;
      break; 
    } else if (responseCode === 503 || responseCode === 500 || responseCode === 429) {
      if (attempt < 5) {
        var waitTime = 4000 * Math.pow(2, attempt - 1);
        Utilities.sleep(waitTime);
      }
    } else {
      break;
    }
  }

  if (!success) {
    throw new Error("Gemini API error (" + responseCode + "): " + responseText);
  }

  var parsed = JSON.parse(responseText);
  return parsed.candidates[0].content.parts[0].text;
}

function build_html_pdf(user_payload, generated_text) {
  // We will replace this URL when we deploy Python to a live server
  var python_api_url = "https://YOUR-LIVE-SERVER-URL.com/api/generate-pdf";

  var payload = {
    intake: user_payload,
    raw_ai_text: generated_text
  };

  var options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(python_api_url, options);
  var responseCode = response.getResponseCode();

  if (responseCode !== 200) {
    throw new Error("Python PDF Engine Failed: " + response.getContentText());
  }

  // The Python server returns the raw PDF file directly
  var blob = response.getBlob();
  
  var safe_name = (user_payload.user_name || "athlete").toLowerCase().replace(/[^a-z0-9]/g, "_");
  blob.setName("workout_program_" + safe_name + ".pdf");

  var archive_folder = DriveApp.getFolderById(FOLDER_ID_ARCHIVE);
  return archive_folder.createFile(blob);
}

function trigger_setup_queue_runner() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "process_workout_queue") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger("process_workout_queue")
    .timeBased()
    .everyMinutes(1)
    .create();
}
// Creates the custom Admin menu when the spreadsheet is opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Admin Tools')
    .addItem('🔄 Rerun Failed Workouts', 'force_retry_errors')
    .addToUi();
}

// Scans the sheet for errors, resets them, and triggers the AI
function force_retry_errors() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var dataRange = sheet.getDataRange();
  var rows = dataRange.getValues();
  var ui = SpreadsheetApp.getUi();
  
  var resetCount = 0;

  // Loop through all rows starting after the header
  for (var i = 1; i < rows.length; i++) {
    var status = rows[i][33]; // Column AH (Index 33)
    
    // Check if the status contains "ERROR"
    if (typeof status === 'string' && status.indexOf("ERROR") !== -1) {
      // Reset status to PENDING
      sheet.getRange(i + 1, 34).setValue("PENDING");
      // Clear out any half-generated garbage in the workout or PDF columns
      sheet.getRange(i + 1, 32).clearContent(); 
      sheet.getRange(i + 1, 33).clearContent();
      resetCount++;
    }
  }

  if (resetCount > 0) {
    SpreadsheetApp.flush(); // Lock in the new PENDING statuses immediately
    
    // Notify the admin
    ui.alert(
      'Queue Reset', 
      resetCount + ' failed request(s) reset to PENDING.\n\nThe system will now attempt to process them.', 
      ui.ButtonSet.OK
    );
    
    // Instantly fire the processing function we wrote earlier
    process_workout_queue(); 
    
  } else {
    ui.alert('Queue Clear', 'No failed requests found in Column AH.', ui.ButtonSet.OK);
  }
}
// Builds the custom Admin menu in the Google Sheet toolbar
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Admin Tools')
    .addItem('🔄 Rerun Failed Workouts (Full API)', 'force_retry_errors')
    .addItem('📄 Regenerate PDF for Selected Row (No API)', 'regenerate_active_row_pdf')
    .addToUi();
}

// Rebuilds the PDF using existing data in the active row, bypassing the AI completely
function regenerate_active_row_pdf() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var row = sheet.getActiveCell().getRow();
  var ui = SpreadsheetApp.getUi();
  
  if (row < 2) {
    ui.alert("Select a valid user row first.");
    return;
  }
  
  // Extract all data for the highlighted row
  var rowData = sheet.getRange(row, 1, 1, 34).getValues()[0];
  var generatedText = rowData[31]; // Column AF (Index 31)
  
  if (!generatedText) {
    ui.alert("Error", "No generated workout text found in Column AF. You must run the AI at least once for this row.", ui.ButtonSet.OK);
    return;
  }
  
  // CORRECTED TOAST CALL
  SpreadsheetApp.getActiveSpreadsheet().toast("Rebuilding PDF...", "Admin Tools", 3);
  
  // Package the specific variables the PDF template requires, falling back to a default if the row is blank
  var user_payload = {
    user_name: rowData[1] || "Athlete",
    preferred_days: rowData[17] || "3",
    primary_goal: rowData[14] || "Strength",
    preferred_split: rowData[20] || "Full Body"
  };
  
  try {
    // Pass the existing AI text into your HTML builder
    var pdf_file = build_html_pdf(user_payload, generatedText);
    
    // Log the new PDF URL into Column AG
    sheet.getRange(row, 33).setValue(pdf_file.getUrl());
    ui.alert("Success", "PDF successfully regenerated and linked in Column AG.", ui.ButtonSet.OK);
  } catch(e) {
    ui.alert("PDF Generation Error", e.toString(), ui.ButtonSet.OK);
  }
}