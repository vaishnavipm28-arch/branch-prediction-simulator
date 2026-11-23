import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import tkinter.font as tkfont
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
import csv
import os
import re
import sys
import time
import shutil
import hashlib
import tempfile
import threading
import importlib.util

# Fonts that can show Tamil / Hindi in graphs (first one found is used)
matplotlib.rcParams["font.family"] = [
    "Nirmala UI", "Latha", "Mangal", "Noto Sans Tamil",
    "Noto Sans Devanagari", "DejaVu Sans"
]

IS_WIN = sys.platform.startswith("win")

# Internal (always English) method names - used for logic, CSV and PDF
METHODS = [
    "Always Taken",
    "Always Not Taken",
    "Repeat Last Outcome",
    "2-Bit Saturating Counter"
]

COLORS = ["#2878B5", "#249B65", "#E89535", "#8064B4"]

DEFAULT_INPUT = (
    "TAKEN, NOT TAKEN, TAKEN, TAKEN, "
    "NOT TAKEN, TAKEN, NOT TAKEN, TAKEN"
)

BG = "#F2F5F9"
BLUE = "#174A7E"
TEXT = "#183153"

LANGS = {"English": "en", "தமிழ்": "ta", "हिन्दी": "hi"}
LANG_NAMES = {"en": "English", "ta": "Tamil", "hi": "Hindi"}

dark_mode = False
lang = "en"
FONT = "Arial"

last_report = []
last_accuracies = []
last_details = []
accuracy_data = {}
results = {}
best_methods_now = []
row_info = {}
bar_objects = []


# =====================================================================
#                          TRANSLATIONS
# =====================================================================

T = {}

T["en"] = {
    "title": "BRANCH PREDICTION SIMULATOR",
    "tagline": "Learn • Predict • Compare",
    "welcome_title": "Welcome!",
    "welcome_text": (
        "Explore how computers predict the next step in a program. "
        "Try different outcomes, compare predictions and understand "
        "the results easily. No previous computer knowledge is required. "
        "Move your mouse over anything and the tool will explain it by voice."
    ),
    "lang_label": "Language:",
    "btn_about": "Why & Who?",
    "btn_how": "How to Use?",
    "btn_formula": "Accuracy Formula",
    "btn_dark": "Dark Mode",
    "btn_light": "Light Mode",
    "btn_voice_on": "Voice: ON",
    "btn_voice_off": "Voice: OFF",
    "btn_explain": "Explain Methods",
    "btn_read": "Read Results",
    "btn_stop": "Stop Voice",
    "input_label": "Enter Actual Outcomes (comma-separated):",
    "btn_run": "Run Simulation",
    "btn_sample": "Sample Input",
    "btn_reset": "Reset",
    "btn_steps": "Step-by-Step",
    "btn_pdf": "Download PDF",
    "btn_csv": "Download CSV",
    "table_title": "Simulation Results - All Methods",
    "graph_lbl": "Accuracy Comparison Graph",
    "col_method": "Method",
    "col_no": "No",
    "col_actual": "Actual",
    "col_pred": "Prediction",
    "col_status": "Status",
    "taken": "TAKEN",
    "nottaken": "NOT TAKEN",
    "correct": "Correct",
    "wrong": "Wrong",
    "m0": "Always Taken",
    "m1": "Always Not Taken",
    "m2": "Repeat Last Outcome",
    "d0": ("This method always guesses that the branch will be taken, no matter "
           "what happened before. It works well when most branches are taken, "
           "like in loops."),
    "d1": ("This method always guesses that the branch will not be taken. "
           "It works well when most branches are not taken."),
    "d2": ("This method guesses that the branch will do the same thing as the "
           "last time. The very first guess starts with taken. It works well "
           "when outcomes come in long streaks."),
    "ph_result": "Enter outcomes and click Run Simulation.",
    "ph_best": "Best method will appear after simulation.",
    "ph_expl": "Prediction explanations will appear here.",
    "ph_caption": "Voice captions: hover over anything to hear and read its explanation.",
    "res_head": "ACCURACY RESULTS",
    "best_one": "BEST METHOD: {m} ({a:.1f}%)",
    "best_tie": "BEST METHODS (TIE): {m} ({a:.1f}%)",
    "expl_head": "PREDICTION SUMMARY",
    "expl_line": "{m}: {c} correct and {w} wrong.",
    "best_tag": " - BEST",
    "graph_t": "Branch Prediction Accuracy Comparison",
    "graph_y": "Accuracy (%)",
    "r_t": "This method always predicts TAKEN.",
    "r_nt": "This method always predicts NOT TAKEN.",
    "r_first": "First prediction starts with TAKEN.",
    "r_prev": "It predicts the previous actual outcome: {prev}.",
    "row_voice": ("Step {n}. The actual outcome was {a}. The method predicted {p}. "
                  "{reason} So this prediction is {s}."),
    "head_voice": ("{m}. {d} It got {c} correct and {w} wrong, so the accuracy is "
                   "{a:.1f} percent."),
    "best_voice": " This is the best method.",
    "read_intro": "Here are the simulation results.",
    "no_results_voice": "There are no results yet. Please click Run Simulation first.",
    "explain_intro": "There are four prediction methods in this tool.",
    "lang_changed": "Language changed to English.",
    "voice_enabled": "Voice is now on. Move your mouse over anything to hear it explained.",
    "steps_title": "Step-by-Step Prediction Explanation",
    "steps_head": "HOW EACH PREDICTION WORKED",
    "step_n": "Step {n}",
    "lbl_actual": "Actual Outcome",
    "lbl_pred": "Prediction",
    "lbl_reason": "Reason",
    "lbl_result": "Result",
    "btn_read_aloud": "Read Aloud",
    "w_input_t": "Input Required",
    "w_input": "Please enter TAKEN or NOT TAKEN.",
    "e_invalid_t": "Invalid Input",
    "e_invalid": "Invalid outcome: {x}\nPlease use TAKEN or NOT TAKEN.",
    "w_nores_t": "No Results",
    "w_nores": "Please run the simulation first.",
    "w_nores_dl": "Please run the simulation before downloading.",
    "sample_t": "Sample Input",
    "sample_m": ("Example outcomes added successfully.\n\n"
                 "Click Run Simulation to see the results."),
    "save_csv_t": "Save Simulation Results",
    "save_pdf_t": "Save Simulation Report",
    "csv_saved_t": "CSV Saved",
    "csv_saved": "CSV report saved successfully.\nYou can open it using Microsoft Excel.",
    "pdf_saved_t": "Report Saved",
    "pdf_saved": "PDF report saved successfully.",
    "lib_t": "Library Required",
    "lib_m": "Install ReportLab using:\npy -m pip install reportlab",
    "err_t": "Error",
    "about_t": "About This Application",
    "about_m": (
        "BRANCH PREDICTION SIMULATOR\n\n"
        "WHAT IS THIS APPLICATION?\n"
        "A simple application that demonstrates how a computer "
        "predicts the next step in a program.\n\n"
        "WHY IS IT USED?\n"
        "To explore different prediction methods and compare "
        "their accuracy.\n\n"
        "WHO CAN USE IT?\n"
        "Anyone interested in learning how computers work. "
        "No previous knowledge is required.\n\n"
        "TECHNOLOGIES:\n"
        "Python, Tkinter and Matplotlib. Voice: gTTS and Windows speech."
    ),
    "how_t": "How to Use",
    "how_m": (
        "GETTING STARTED\n\n"
        "1. Choose your language (English, Tamil or Hindi).\n\n"
        "2. Enter TAKEN or NOT TAKEN outcomes.\n\n"
        "3. Click Run Simulation.\n\n"
        "4. Check which predictions are correct or wrong.\n\n"
        "5. Move the mouse over any table row to hear why it is correct or wrong.\n\n"
        "6. Move the mouse over a graph bar to hear about that method.\n\n"
        "7. Use Explain Methods, Read Results and Step-by-Step to learn more.\n\n"
        "8. Use Voice ON/OFF to control the voice.\n\n"
        "9. Download PDF or CSV reports. Click Reset to start again.\n\n"
        "TIP: Move your mouse over any button to hear what it does."
    ),
    "formula_t": "Understanding Accuracy",
    "formula_m": (
        "ACCURACY FORMULA\n\n"
        "Accuracy = (Correct Predictions / Total Predictions) x 100\n\n"
        "SIMPLE EXAMPLE\n\n"
        "Total Predictions = 8\n"
        "Correct Predictions = 5\n\n"
        "Accuracy = (5 / 8) x 100\n"
        "Accuracy = 62.5%\n\n"
        "This means 5 out of 8 predictions were correct."
    ),
    "v_status_gtts": "Voice engine: Online (gTTS)",
    "v_status_sys": "Voice engine: System voice",
    "v_status_none": "Voice engine: none found",
    "v_none": ("No voice engine found. Install one with:  "
               "pip install gTTS"),
    "v_nolang": ("No {lang} voice is installed on this computer. Connect to the "
                 "internet (gTTS) or install a {lang} voice pack in Windows."),
    "tip_header": "This is the Branch Prediction Simulator. It shows how a computer predicts the next step of a program.",
    "tip_about": "Understand why this application is useful and who can use it.",
    "tip_how": "Read simple instructions to use the application.",
    "tip_formula": "Understand how prediction accuracy is calculated.",
    "tip_theme": "Change between light and dark appearance.",
    "tip_voice": "Turn the hover voice on or off.",
    "tip_explain": "Hear a voice explanation of all four prediction methods.",
    "tip_read": "Hear the simulation results read aloud.",
    "tip_stop": "Stop the voice immediately.",
    "tip_lang": "Choose English, Tamil or Hindi. The voice and the whole application change language.",
    "tip_input": "Enter TAKEN or NOT TAKEN outcomes separated by commas.",
    "tip_run": "Start the simulation and compare all four methods.",
    "tip_sample": "Add example outcomes to try the application.",
    "tip_reset": "Restore default input and clear previous results.",
    "tip_steps": "Understand why each prediction was made.",
    "tip_pdf": "Save the simulation results as a PDF document.",
    "tip_csv": "Save results in an Excel-compatible CSV file.",
    "tip_table": "This table lists every prediction. Move the mouse over a row to hear why it was correct or wrong.",
    "tip_graph": "This graph compares the accuracy of the four methods. Move the mouse over a bar to hear about that method.",
    "tip_result": "These are the accuracy percentages of the four methods.",
    "tip_best": "This shows which method predicted best.",
    "tip_expl": "This summary shows how many predictions each method got right and wrong.",
    "tip_caption": "This line shows the text that is being spoken.",
}

T["ta"] = {
    "title": "கிளை முன்கணிப்பு சிமுலேட்டர்",
    "tagline": "கற்றுக்கொள் • கணி • ஒப்பிடு",
    "welcome_title": "வரவேற்கிறோம்!",
    "welcome_text": (
        "ஒரு நிரலில் அடுத்த படியை கணினி எப்படி முன்கணிக்கிறது என்பதை ஆராயுங்கள். "
        "வெவ்வேறு முடிவுகளை உள்ளிடுங்கள், கணிப்புகளை ஒப்பிடுங்கள், முடிவுகளை "
        "எளிதாகப் புரிந்துகொள்ளுங்கள். கணினி பற்றிய முன் அறிவு தேவையில்லை. "
        "எதன் மீதாவது மவுஸை நகர்த்தினால் அதன் விளக்கத்தைக் குரலில் கேட்கலாம்."
    ),
    "lang_label": "மொழி:",
    "btn_about": "ஏன் & யார்?",
    "btn_how": "எப்படி பயன்படுத்துவது?",
    "btn_formula": "துல்லிய சூத்திரம்",
    "btn_dark": "இருள் பயன்முறை",
    "btn_light": "ஒளி பயன்முறை",
    "btn_voice_on": "குரல்: ஆன்",
    "btn_voice_off": "குரல்: ஆஃப்",
    "btn_explain": "முறைகளை விளக்கு",
    "btn_read": "முடிவுகளை வாசி",
    "btn_stop": "குரலை நிறுத்து",
    "input_label": "உண்மையான முடிவுகளை உள்ளிடுங்கள் (கமாவால் பிரிக்கவும்):",
    "btn_run": "சிமுலேஷனை இயக்கு",
    "btn_sample": "மாதிரி உள்ளீடு",
    "btn_reset": "மீட்டமை",
    "btn_steps": "படிப்படியாக",
    "btn_pdf": "PDF பதிவிறக்கு",
    "btn_csv": "CSV பதிவிறக்கு",
    "table_title": "சிமுலேஷன் முடிவுகள் - அனைத்து முறைகள்",
    "graph_lbl": "துல்லிய ஒப்பீட்டு வரைபடம்",
    "col_method": "முறை",
    "col_no": "எண்",
    "col_actual": "உண்மை",
    "col_pred": "கணிப்பு",
    "col_status": "நிலை",
    "taken": "எடுக்கப்பட்டது",
    "nottaken": "எடுக்கப்படவில்லை",
    "correct": "சரி",
    "wrong": "தவறு",
    "m0": "எப்போதும் எடுக்கப்படும்",
    "m1": "எப்போதும் எடுக்கப்படாது",
    "m2": "முந்தைய முடிவை மீண்டும் கணி",
    "d0": ("இந்த முறை முன்பு என்ன நடந்தாலும், கிளை எப்போதும் எடுக்கப்படும் என்றே "
           "கணிக்கும். பெரும்பாலான கிளைகள் எடுக்கப்படும் இடங்களில், உதாரணமாக "
           "சுழற்சிகளில், இது நன்றாக வேலை செய்யும்."),
    "d1": ("இந்த முறை கிளை எப்போதும் எடுக்கப்படாது என்றே கணிக்கும். "
           "பெரும்பாலான கிளைகள் எடுக்கப்படாத இடங்களில் இது நன்றாக வேலை செய்யும்."),
    "d2": ("இந்த முறை கடந்த முறை என்ன நடந்ததோ அதையே இந்த முறையும் கணிக்கும். "
           "முதல் கணிப்பு எடுக்கப்பட்டது என்பதில் தொடங்கும். முடிவுகள் நீண்ட "
           "தொடராக வரும்போது இது நன்றாக வேலை செய்யும்."),
    "ph_result": "முடிவுகளை உள்ளிட்டு சிமுலேஷனை இயக்கு என்பதை அழுத்துங்கள்.",
    "ph_best": "சிமுலேஷனுக்குப் பிறகு சிறந்த முறை இங்கே தோன்றும்.",
    "ph_expl": "கணிப்பு விளக்கங்கள் இங்கே தோன்றும்.",
    "ph_caption": "குரல் வசனம்: எதன் மீதாவது மவுஸை நகர்த்தினால் விளக்கத்தைக் கேட்கலாம், படிக்கலாம்.",
    "res_head": "துல்லிய முடிவுகள்",
    "best_one": "சிறந்த முறை: {m} ({a:.1f}%)",
    "best_tie": "சிறந்த முறைகள் (சமம்): {m} ({a:.1f}%)",
    "expl_head": "கணிப்பு சுருக்கம்",
    "expl_line": "{m}: {c} சரி, {w} தவறு.",
    "best_tag": " - சிறந்தது",
    "graph_t": "கிளை முன்கணிப்பு துல்லிய ஒப்பீடு",
    "graph_y": "துல்லியம் (%)",
    "r_t": "இந்த முறை எப்போதும் எடுக்கப்பட்டது என்றே கணிக்கும்.",
    "r_nt": "இந்த முறை எப்போதும் எடுக்கப்படவில்லை என்றே கணிக்கும்.",
    "r_first": "முதல் கணிப்பு எடுக்கப்பட்டது என்பதில் தொடங்குகிறது.",
    "r_prev": "இது முந்தைய உண்மையான முடிவையே கணிக்கிறது: {prev}.",
    "row_voice": ("படி {n}. உண்மையான முடிவு {a}. இந்த முறை கணித்தது {p}. "
                  "{reason} எனவே இந்த கணிப்பு {s}."),
    "head_voice": ("{m}. {d} இது {c} சரியாகவும் {w} தவறாகவும் கணித்தது, "
                   "எனவே துல்லியம் {a:.1f} சதவீதம்."),
    "best_voice": " இதுவே சிறந்த முறை.",
    "read_intro": "இதோ சிமுலேஷன் முடிவுகள்.",
    "no_results_voice": "இன்னும் முடிவுகள் இல்லை. முதலில் சிமுலேஷனை இயக்கு என்பதை அழுத்துங்கள்.",
    "explain_intro": "இந்த கருவியில் நான்கு கணிப்பு முறைகள் உள்ளன.",
    "lang_changed": "மொழி தமிழுக்கு மாற்றப்பட்டது.",
    "voice_enabled": "குரல் இப்போது ஆன். எதன் மீதாவது மவுஸை நகர்த்தினால் விளக்கம் கேட்கும்.",
    "steps_title": "படிப்படியான கணிப்பு விளக்கம்",
    "steps_head": "ஒவ்வொரு கணிப்பும் எப்படி வேலை செய்தது",
    "step_n": "படி {n}",
    "lbl_actual": "உண்மையான முடிவு",
    "lbl_pred": "கணிப்பு",
    "lbl_reason": "காரணம்",
    "lbl_result": "முடிவு",
    "btn_read_aloud": "உரக்கப் படி",
    "w_input_t": "உள்ளீடு தேவை",
    "w_input": "தயவுசெய்து TAKEN அல்லது NOT TAKEN உள்ளிடுங்கள்.",
    "e_invalid_t": "தவறான உள்ளீடு",
    "e_invalid": "தவறான முடிவு: {x}\nதயவுசெய்து TAKEN அல்லது NOT TAKEN பயன்படுத்துங்கள்.",
    "w_nores_t": "முடிவுகள் இல்லை",
    "w_nores": "முதலில் சிமுலேஷனை இயக்குங்கள்.",
    "w_nores_dl": "பதிவிறக்கும் முன் சிமுலேஷனை இயக்குங்கள்.",
    "sample_t": "மாதிரி உள்ளீடு",
    "sample_m": ("மாதிரி முடிவுகள் சேர்க்கப்பட்டன.\n\n"
                 "முடிவுகளைப் பார்க்க சிமுலேஷனை இயக்கு என்பதை அழுத்துங்கள்."),
    "save_csv_t": "சிமுலேஷன் முடிவுகளைச் சேமி",
    "save_pdf_t": "சிமுலேஷன் அறிக்கையைச் சேமி",
    "csv_saved_t": "CSV சேமிக்கப்பட்டது",
    "csv_saved": "CSV அறிக்கை வெற்றிகரமாகச் சேமிக்கப்பட்டது.\nMicrosoft Excel-இல் திறக்கலாம்.",
    "pdf_saved_t": "அறிக்கை சேமிக்கப்பட்டது",
    "pdf_saved": "PDF அறிக்கை வெற்றிகரமாகச் சேமிக்கப்பட்டது.",
    "lib_t": "நூலகம் தேவை",
    "lib_m": "ReportLab-ஐ இவ்வாறு நிறுவவும்:\npy -m pip install reportlab",
    "err_t": "பிழை",
    "about_t": "இந்த பயன்பாட்டைப் பற்றி",
    "about_m": (
        "கிளை முன்கணிப்பு சிமுலேட்டர்\n\n"
        "இது என்ன?\n"
        "ஒரு நிரலில் அடுத்த படியை கணினி எப்படி முன்கணிக்கிறது என்பதைக் காட்டும் "
        "எளிய பயன்பாடு.\n\n"
        "ஏன் பயன்படுகிறது?\n"
        "வெவ்வேறு கணிப்பு முறைகளை ஆராய்ந்து அவற்றின் துல்லியத்தை ஒப்பிட.\n\n"
        "யார் பயன்படுத்தலாம்?\n"
        "கணினி எப்படி வேலை செய்கிறது என்பதைக் கற்க விரும்பும் எவரும். "
        "முன் அறிவு தேவையில்லை.\n\n"
        "தொழில்நுட்பங்கள்:\n"
        "Python, Tkinter, Matplotlib. குரல்: gTTS மற்றும் Windows குரல்."
    ),
    "how_t": "எப்படி பயன்படுத்துவது",
    "how_m": (
        "தொடங்குவது எப்படி\n\n"
        "1. உங்கள் மொழியைத் தேர்ந்தெடுங்கள் (ஆங்கிலம், தமிழ் அல்லது இந்தி).\n\n"
        "2. TAKEN அல்லது NOT TAKEN முடிவுகளை உள்ளிடுங்கள்.\n\n"
        "3. சிமுலேஷனை இயக்கு என்பதை அழுத்துங்கள்.\n\n"
        "4. எந்த கணிப்புகள் சரி அல்லது தவறு என்று பாருங்கள்.\n\n"
        "5. அட்டவணையின் எந்த வரிசையின் மீதும் மவுஸை நகர்த்தினால் அது ஏன் சரி அல்லது தவறு என்று கேட்கலாம்.\n\n"
        "6. வரைபடத்தின் பட்டையின் மீது மவுஸை நகர்த்தினால் அந்த முறையைப் பற்றிக் கேட்கலாம்.\n\n"
        "7. மேலும் அறிய முறைகளை விளக்கு, முடிவுகளை வாசி, படிப்படியாக பயன்படுத்துங்கள்.\n\n"
        "8. குரலைக் கட்டுப்படுத்த குரல் ஆன்/ஆஃப் பயன்படுத்துங்கள்.\n\n"
        "9. PDF அல்லது CSV பதிவிறக்கலாம். மீண்டும் தொடங்க மீட்டமை அழுத்துங்கள்.\n\n"
        "குறிப்பு: எந்த பொத்தானின் மீதும் மவுஸை நகர்த்தினால் அது என்ன செய்யும் என்று கேட்கலாம்."
    ),
    "formula_t": "துல்லியத்தைப் புரிந்துகொள்ளுதல்",
    "formula_m": (
        "துல்லிய சூத்திரம்\n\n"
        "துல்லியம் = (சரியான கணிப்புகள் / மொத்த கணிப்புகள்) x 100\n\n"
        "எளிய உதாரணம்\n\n"
        "மொத்த கணிப்புகள் = 8\n"
        "சரியான கணிப்புகள் = 5\n\n"
        "துல்லியம் = (5 / 8) x 100\n"
        "துல்லியம் = 62.5%\n\n"
        "அதாவது 8 கணிப்புகளில் 5 சரியாக இருந்தன."
    ),
    "v_status_gtts": "குரல் இயந்திரம்: ஆன்லைன் (gTTS)",
    "v_status_sys": "குரல் இயந்திரம்: கணினி குரல்",
    "v_status_none": "குரல் இயந்திரம் கிடைக்கவில்லை",
    "v_none": ("குரல் இயந்திரம் கிடைக்கவில்லை. இதை நிறுவவும்:  "
               "pip install gTTS"),
    "v_nolang": ("இந்த கணினியில் {lang} குரல் நிறுவப்படவில்லை. இணையத்துடன் இணையுங்கள் "
                 "(gTTS) அல்லது Windows-இல் {lang} குரல் தொகுப்பை நிறுவுங்கள்."),
    "tip_header": "இது கிளை முன்கணிப்பு சிமுலேட்டர். ஒரு நிரலின் அடுத்த படியை கணினி எப்படி முன்கணிக்கிறது என்பதை இது காட்டுகிறது.",
    "tip_about": "இந்த பயன்பாடு ஏன் பயனுள்ளது, யார் பயன்படுத்தலாம் என்பதைப் புரிந்துகொள்ளுங்கள்.",
    "tip_how": "பயன்பாட்டைப் பயன்படுத்த எளிய வழிமுறைகளைப் படியுங்கள்.",
    "tip_formula": "கணிப்பு துல்லியம் எப்படி கணக்கிடப்படுகிறது என்பதைப் புரிந்துகொள்ளுங்கள்.",
    "tip_theme": "ஒளி மற்றும் இருள் தோற்றத்திற்கு இடையே மாற்றுங்கள்.",
    "tip_voice": "மவுஸ் நகர்த்தும்போது வரும் குரலை ஆன் அல்லது ஆஃப் செய்யுங்கள்.",
    "tip_explain": "நான்கு கணிப்பு முறைகளின் குரல் விளக்கத்தைக் கேளுங்கள்.",
    "tip_read": "சிமுலேஷன் முடிவுகளை உரக்கக் கேளுங்கள்.",
    "tip_stop": "குரலை உடனே நிறுத்துங்கள்.",
    "tip_lang": "ஆங்கிலம், தமிழ் அல்லது இந்தியைத் தேர்ந்தெடுங்கள். குரலும் முழு பயன்பாடும் அந்த மொழிக்கு மாறும்.",
    "tip_input": "TAKEN அல்லது NOT TAKEN முடிவுகளை கமாவால் பிரித்து உள்ளிடுங்கள்.",
    "tip_run": "சிமுலேஷனைத் தொடங்கி நான்கு முறைகளையும் ஒப்பிடுங்கள்.",
    "tip_sample": "பயன்பாட்டை முயற்சிக்க மாதிரி முடிவுகளைச் சேர்க்கவும்.",
    "tip_reset": "இயல்புநிலை உள்ளீட்டை மீட்டு முந்தைய முடிவுகளை அழிக்கவும்.",
    "tip_steps": "ஒவ்வொரு கணிப்பும் ஏன் செய்யப்பட்டது என்பதைப் புரிந்துகொள்ளுங்கள்.",
    "tip_pdf": "சிமுலேஷன் முடிவுகளை PDF ஆவணமாகச் சேமிக்கவும்.",
    "tip_csv": "முடிவுகளை Excel-இல் திறக்கக்கூடிய CSV கோப்பாகச் சேமிக்கவும்.",
    "tip_table": "இந்த அட்டவணை ஒவ்வொரு கணிப்பையும் காட்டுகிறது. ஒரு வரிசையின் மீது மவுஸை நகர்த்தினால் அது ஏன் சரி அல்லது தவறு என்று கேட்கலாம்.",
    "tip_graph": "இந்த வரைபடம் நான்கு முறைகளின் துல்லியத்தை ஒப்பிடுகிறது. ஒரு பட்டையின் மீது மவுஸை நகர்த்தினால் அந்த முறையைப் பற்றிக் கேட்கலாம்.",
    "tip_result": "இவை நான்கு முறைகளின் துல்லிய சதவீதங்கள்.",
    "tip_best": "எந்த முறை சிறப்பாகக் கணித்தது என்பதை இது காட்டுகிறது.",
    "tip_expl": "ஒவ்வொரு முறையும் எத்தனை கணிப்புகளைச் சரியாகவும் தவறாகவும் செய்தது என்பதை இந்த சுருக்கம் காட்டுகிறது.",
    "tip_caption": "பேசப்படும் உரையை இந்த வரி காட்டுகிறது.",
}

T["hi"] = {
    "title": "ब्रांच प्रेडिक्शन सिम्युलेटर",
    "tagline": "सीखें • अनुमान लगाएँ • तुलना करें",
    "welcome_title": "स्वागत है!",
    "welcome_text": (
        "जानिए कि कंप्यूटर किसी प्रोग्राम में अगले चरण का अनुमान कैसे लगाता है। "
        "अलग-अलग परिणाम दर्ज करें, अनुमानों की तुलना करें और नतीजों को आसानी से "
        "समझें। कंप्यूटर का पहले से कोई ज्ञान ज़रूरी नहीं है। "
        "किसी भी चीज़ पर माउस ले जाएँ, यह टूल आवाज़ में उसे समझाएगा।"
    ),
    "lang_label": "भाषा:",
    "btn_about": "क्यों और कौन?",
    "btn_how": "कैसे उपयोग करें?",
    "btn_formula": "सटीकता सूत्र",
    "btn_dark": "डार्क मोड",
    "btn_light": "लाइट मोड",
    "btn_voice_on": "आवाज़: चालू",
    "btn_voice_off": "आवाज़: बंद",
    "btn_explain": "तरीके समझाएँ",
    "btn_read": "परिणाम सुनाएँ",
    "btn_stop": "आवाज़ रोकें",
    "input_label": "वास्तविक परिणाम दर्ज करें (कॉमा से अलग करें):",
    "btn_run": "सिमुलेशन चलाएँ",
    "btn_sample": "नमूना इनपुट",
    "btn_reset": "रीसेट",
    "btn_steps": "चरण-दर-चरण",
    "btn_pdf": "PDF डाउनलोड",
    "btn_csv": "CSV डाउनलोड",
    "table_title": "सिमुलेशन परिणाम - सभी तरीके",
    "graph_lbl": "सटीकता तुलना ग्राफ़",
    "col_method": "तरीका",
    "col_no": "क्र.",
    "col_actual": "वास्तविक",
    "col_pred": "अनुमान",
    "col_status": "स्थिति",
    "taken": "लिया गया",
    "nottaken": "नहीं लिया गया",
    "correct": "सही",
    "wrong": "गलत",
    "m0": "हमेशा लिया गया",
    "m1": "हमेशा नहीं लिया गया",
    "m2": "पिछला परिणाम दोहराएँ",
    "d0": ("यह तरीका पहले जो भी हुआ हो, हमेशा यही अनुमान लगाता है कि ब्रांच ली जाएगी। "
           "जब ज़्यादातर ब्रांच ली जाती हैं, जैसे लूप में, तब यह अच्छा काम करता है।"),
    "d1": ("यह तरीका हमेशा यही अनुमान लगाता है कि ब्रांच नहीं ली जाएगी। "
           "जब ज़्यादातर ब्रांच नहीं ली जातीं, तब यह अच्छा काम करता है।"),
    "d2": ("यह तरीका अनुमान लगाता है कि इस बार वही होगा जो पिछली बार हुआ था। "
           "पहला अनुमान लिया गया से शुरू होता है। जब परिणाम लंबी श्रृंखला में आते हैं, "
           "तब यह अच्छा काम करता है।"),
    "ph_result": "परिणाम दर्ज करें और सिमुलेशन चलाएँ पर क्लिक करें।",
    "ph_best": "सिमुलेशन के बाद सबसे अच्छा तरीका यहाँ दिखेगा।",
    "ph_expl": "अनुमानों की व्याख्या यहाँ दिखेगी।",
    "ph_caption": "आवाज़ कैप्शन: किसी भी चीज़ पर माउस ले जाएँ, उसकी व्याख्या सुनें और पढ़ें।",
    "res_head": "सटीकता परिणाम",
    "best_one": "सबसे अच्छा तरीका: {m} ({a:.1f}%)",
    "best_tie": "सबसे अच्छे तरीके (बराबरी): {m} ({a:.1f}%)",
    "expl_head": "अनुमान सारांश",
    "expl_line": "{m}: {c} सही और {w} गलत।",
    "best_tag": " - सर्वश्रेष्ठ",
    "graph_t": "ब्रांच प्रेडिक्शन सटीकता तुलना",
    "graph_y": "सटीकता (%)",
    "r_t": "यह तरीका हमेशा लिया गया का अनुमान लगाता है।",
    "r_nt": "यह तरीका हमेशा नहीं लिया गया का अनुमान लगाता है।",
    "r_first": "पहला अनुमान लिया गया से शुरू होता है।",
    "r_prev": "यह पिछले वास्तविक परिणाम का अनुमान लगाता है: {prev}।",
    "row_voice": ("चरण {n}। वास्तविक परिणाम {a} था। तरीके ने {p} का अनुमान लगाया। "
                  "{reason} इसलिए यह अनुमान {s} है।"),
    "head_voice": ("{m}। {d} इसने {c} सही और {w} गलत अनुमान लगाए, "
                   "इसलिए सटीकता {a:.1f} प्रतिशत है।"),
    "best_voice": " यह सबसे अच्छा तरीका है।",
    "read_intro": "ये रहे सिमुलेशन के परिणाम।",
    "no_results_voice": "अभी कोई परिणाम नहीं है। पहले सिमुलेशन चलाएँ पर क्लिक करें।",
    "explain_intro": "इस टूल में चार अनुमान तरीके हैं।",
    "lang_changed": "भाषा हिन्दी में बदल दी गई है।",
    "voice_enabled": "आवाज़ अब चालू है। किसी भी चीज़ पर माउस ले जाएँ और उसकी व्याख्या सुनें।",
    "steps_title": "चरण-दर-चरण अनुमान व्याख्या",
    "steps_head": "हर अनुमान कैसे काम किया",
    "step_n": "चरण {n}",
    "lbl_actual": "वास्तविक परिणाम",
    "lbl_pred": "अनुमान",
    "lbl_reason": "कारण",
    "lbl_result": "नतीजा",
    "btn_read_aloud": "ज़ोर से पढ़ें",
    "w_input_t": "इनपुट ज़रूरी है",
    "w_input": "कृपया TAKEN या NOT TAKEN दर्ज करें।",
    "e_invalid_t": "गलत इनपुट",
    "e_invalid": "गलत परिणाम: {x}\nकृपया TAKEN या NOT TAKEN का उपयोग करें।",
    "w_nores_t": "कोई परिणाम नहीं",
    "w_nores": "कृपया पहले सिमुलेशन चलाएँ।",
    "w_nores_dl": "डाउनलोड करने से पहले सिमुलेशन चलाएँ।",
    "sample_t": "नमूना इनपुट",
    "sample_m": ("उदाहरण परिणाम जोड़ दिए गए हैं।\n\n"
                 "परिणाम देखने के लिए सिमुलेशन चलाएँ पर क्लिक करें।"),
    "save_csv_t": "सिमुलेशन परिणाम सहेजें",
    "save_pdf_t": "सिमुलेशन रिपोर्ट सहेजें",
    "csv_saved_t": "CSV सहेजी गई",
    "csv_saved": "CSV रिपोर्ट सफलतापूर्वक सहेजी गई।\nआप इसे Microsoft Excel में खोल सकते हैं।",
    "pdf_saved_t": "रिपोर्ट सहेजी गई",
    "pdf_saved": "PDF रिपोर्ट सफलतापूर्वक सहेजी गई।",
    "lib_t": "लाइब्रेरी आवश्यक",
    "lib_m": "ReportLab इस तरह इंस्टॉल करें:\npy -m pip install reportlab",
    "err_t": "त्रुटि",
    "about_t": "इस एप्लिकेशन के बारे में",
    "about_m": (
        "ब्रांच प्रेडिक्शन सिम्युलेटर\n\n"
        "यह क्या है?\n"
        "एक सरल एप्लिकेशन जो दिखाता है कि कंप्यूटर किसी प्रोग्राम में अगले चरण का "
        "अनुमान कैसे लगाता है।\n\n"
        "यह क्यों उपयोगी है?\n"
        "अलग-अलग अनुमान तरीकों को जानने और उनकी सटीकता की तुलना करने के लिए।\n\n"
        "इसे कौन उपयोग कर सकता है?\n"
        "कोई भी जो जानना चाहता है कि कंप्यूटर कैसे काम करते हैं। "
        "पहले से ज्ञान ज़रूरी नहीं है।\n\n"
        "तकनीकें:\n"
        "Python, Tkinter और Matplotlib। आवाज़: gTTS और Windows आवाज़।"
    ),
    "how_t": "कैसे उपयोग करें",
    "how_m": (
        "शुरुआत कैसे करें\n\n"
        "1. अपनी भाषा चुनें (अंग्रेज़ी, तमिल या हिन्दी)।\n\n"
        "2. TAKEN या NOT TAKEN परिणाम दर्ज करें।\n\n"
        "3. सिमुलेशन चलाएँ पर क्लिक करें।\n\n"
        "4. देखें कि कौन से अनुमान सही या गलत हैं।\n\n"
        "5. तालिका की किसी पंक्ति पर माउस ले जाएँ और सुनें कि वह सही या गलत क्यों है।\n\n"
        "6. ग्राफ़ की किसी बार पर माउस ले जाएँ और उस तरीके के बारे में सुनें।\n\n"
        "7. और जानने के लिए तरीके समझाएँ, परिणाम सुनाएँ और चरण-दर-चरण का उपयोग करें।\n\n"
        "8. आवाज़ को नियंत्रित करने के लिए आवाज़ चालू/बंद का उपयोग करें।\n\n"
        "9. PDF या CSV रिपोर्ट डाउनलोड करें। फिर से शुरू करने के लिए रीसेट दबाएँ।\n\n"
        "सुझाव: किसी भी बटन पर माउस ले जाएँ और सुनें कि वह क्या करता है।"
    ),
    "formula_t": "सटीकता को समझना",
    "formula_m": (
        "सटीकता सूत्र\n\n"
        "सटीकता = (सही अनुमान / कुल अनुमान) x 100\n\n"
        "सरल उदाहरण\n\n"
        "कुल अनुमान = 8\n"
        "सही अनुमान = 5\n\n"
        "सटीकता = (5 / 8) x 100\n"
        "सटीकता = 62.5%\n\n"
        "इसका मतलब है कि 8 में से 5 अनुमान सही थे।"
    ),
    "v_status_gtts": "आवाज़ इंजन: ऑनलाइन (gTTS)",
    "v_status_sys": "आवाज़ इंजन: सिस्टम आवाज़",
    "v_status_none": "कोई आवाज़ इंजन नहीं मिला",
    "v_none": ("कोई आवाज़ इंजन नहीं मिला। इसे इंस्टॉल करें:  "
               "pip install gTTS"),
    "v_nolang": ("इस कंप्यूटर में {lang} आवाज़ इंस्टॉल नहीं है। इंटरनेट से जुड़ें "
                 "(gTTS) या Windows में {lang} वॉइस पैक इंस्टॉल करें।"),
    "tip_header": "यह ब्रांच प्रेडिक्शन सिम्युलेटर है। यह दिखाता है कि कंप्यूटर प्रोग्राम के अगले चरण का अनुमान कैसे लगाता है।",
    "tip_about": "समझें कि यह एप्लिकेशन क्यों उपयोगी है और इसे कौन उपयोग कर सकता है।",
    "tip_how": "एप्लिकेशन उपयोग करने के सरल निर्देश पढ़ें।",
    "tip_formula": "समझें कि अनुमान की सटीकता कैसे निकाली जाती है।",
    "tip_theme": "लाइट और डार्क रूप के बीच बदलें।",
    "tip_voice": "माउस ले जाने पर आने वाली आवाज़ को चालू या बंद करें।",
    "tip_explain": "चारों अनुमान तरीकों की आवाज़ में व्याख्या सुनें।",
    "tip_read": "सिमुलेशन के परिणाम ज़ोर से सुनें।",
    "tip_stop": "आवाज़ तुरंत रोकें।",
    "tip_lang": "अंग्रेज़ी, तमिल या हिन्दी चुनें। आवाज़ और पूरा एप्लिकेशन उसी भाषा में बदल जाएगा।",
    "tip_input": "TAKEN या NOT TAKEN परिणाम कॉमा से अलग करके दर्ज करें।",
    "tip_run": "सिमुलेशन शुरू करें और चारों तरीकों की तुलना करें।",
    "tip_sample": "एप्लिकेशन आज़माने के लिए उदाहरण परिणाम जोड़ें।",
    "tip_reset": "डिफ़ॉल्ट इनपुट वापस लाएँ और पिछले परिणाम साफ़ करें।",
    "tip_steps": "समझें कि हर अनुमान क्यों लगाया गया।",
    "tip_pdf": "सिमुलेशन परिणाम PDF दस्तावेज़ के रूप में सहेजें।",
    "tip_csv": "परिणाम Excel में खुलने वाली CSV फ़ाइल में सहेजें।",
    "tip_table": "यह तालिका हर अनुमान दिखाती है। किसी पंक्ति पर माउस ले जाएँ और सुनें कि वह सही या गलत क्यों है।",
    "tip_graph": "यह ग्राफ़ चारों तरीकों की सटीकता की तुलना करता है। किसी बार पर माउस ले जाएँ और उस तरीके के बारे में सुनें।",
    "tip_result": "ये चारों तरीकों के सटीकता प्रतिशत हैं।",
    "tip_best": "यह दिखाता है कि किस तरीके ने सबसे अच्छा अनुमान लगाया।",
    "tip_expl": "यह सारांश दिखाता है कि हर तरीके ने कितने अनुमान सही और कितने गलत लगाए।",
    "tip_caption": "यह पंक्ति वही पाठ दिखाती है जो बोला जा रहा है।",
}


T["en"].update({
    "m3": "2-Bit Saturating Counter",
    "d3": ("This method keeps a small counter from 0 to 3. A high counter means "
           "taken and a low counter means not taken. Each TAKEN adds one and each "
           "NOT TAKEN subtracts one, so a single surprise does not change the "
           "guess right away. It starts at Weakly Taken. It works well for loops."),
    "s0": "Strongly Not Taken",
    "s1": "Weakly Not Taken",
    "s2": "Weakly Taken",
    "s3": "Strongly Taken",
    "r_2bit": ("The counter was {s}, so it predicts {p}. After the actual "
               "outcome {a}, the counter moves to {n}."),
})

T["ta"].update({
    "m3": "2-பிட் கவுண்டர் முறை",
    "d3": ("இந்த முறை 0 முதல் 3 வரை ஒரு சிறிய கவுண்டரை வைத்திருக்கும். கவுண்டர் "
           "அதிகமாக இருந்தால் எடுக்கப்பட்டது என்றும், குறைவாக இருந்தால் "
           "எடுக்கப்படவில்லை என்றும் கணிக்கும். ஒவ்வொரு எடுக்கப்பட்டதற்கும் ஒன்று "
           "கூடும், ஒவ்வொரு எடுக்கப்படவில்லைக்கும் ஒன்று குறையும். அதனால் ஒரு "
           "எதிர்பாராத முடிவு கணிப்பை உடனே மாற்றாது. இது பலவீனமாக எடுக்கப்பட்டது "
           "என்பதில் தொடங்கும். சுழற்சிகளில் இது நன்றாக வேலை செய்யும்."),
    "s0": "உறுதியாக எடுக்கப்படவில்லை",
    "s1": "பலவீனமாக எடுக்கப்படவில்லை",
    "s2": "பலவீனமாக எடுக்கப்பட்டது",
    "s3": "உறுதியாக எடுக்கப்பட்டது",
    "r_2bit": ("கவுண்டர் {s} நிலையில் இருந்தது, எனவே {p} என்று கணிக்கிறது. "
               "உண்மையான முடிவு {a} வந்ததால், கவுண்டர் {n} நிலைக்கு மாறுகிறது."),
})

T["hi"].update({
    "m3": "2-बिट सैचुरेटिंग काउंटर",
    "d3": ("यह तरीका 0 से 3 तक एक छोटा काउंटर रखता है। काउंटर ऊँचा हो तो लिया गया "
           "और नीचा हो तो नहीं लिया गया का अनुमान लगाता है। हर लिया गया पर काउंटर "
           "एक बढ़ता है और हर नहीं लिया गया पर एक घटता है, इसलिए एक अप्रत्याशित "
           "परिणाम अनुमान को तुरंत नहीं बदलता। यह कमज़ोर रूप से लिया गया से शुरू "
           "होता है। लूप में यह अच्छा काम करता है।"),
    "s0": "पक्का नहीं लिया गया",
    "s1": "कमज़ोर नहीं लिया गया",
    "s2": "कमज़ोर लिया गया",
    "s3": "पक्का लिया गया",
    "r_2bit": ("काउंटर {s} स्थिति में था, इसलिए यह {p} का अनुमान लगाता है। "
               "वास्तविक परिणाम {a} आने पर काउंटर {n} स्थिति में चला जाता है।"),
})


T["en"]["btn_close"] = "Close"
T["ta"]["btn_close"] = "மூடு"
T["hi"]["btn_close"] = "बंद करें"

def tr(key, **kw):
    text = T.get(lang, {}).get(key)
    if text is None:
        text = T["en"].get(key, key)
    if kw:
        try:
            return text.format(**kw)
        except Exception:
            return text
    return text


def oname(outcome):
    return tr("taken") if outcome == "TAKEN" else tr("nottaken")


def sname(status):
    return tr("correct") if status == "Correct" else tr("wrong")


def mname(method):
    return tr("m%d" % METHODS.index(method))


def mdesc(method):
    return tr("d%d" % METHODS.index(method))


# =====================================================================
#                          VOICE ENGINE
# =====================================================================

def has_module(name):
    try:
        return importlib.util.find_spec(name) is not None
    except Exception:
        return False


def split_for_speech(text, first_min=40, min_len=70, max_len=230):
    """Cut a long text into sentence groups so speech can start quickly.
    The first group is short, the later ones are a little longer."""
    text = " ".join(text.split())
    sentences = [x for x in re.split(r"(?<=[.!?\u0964])\s+", text) if x]

    chunks = []
    current = ""

    for sentence in sentences:
        if current and len(current) + 1 + len(sentence) > max_len:
            chunks.append(current)
            current = sentence
        else:
            current = (current + " " + sentence).strip()

        need = first_min if not chunks else min_len
        if len(current) >= need:
            chunks.append(current)
            current = ""

    if current:
        if chunks and len(current) < 25:
            chunks[-1] = chunks[-1] + " " + current
        else:
            chunks.append(current)

    return chunks or [text]


class NoVoice(Exception):
    pass


# Offline voice for Windows: built-in SAPI through PowerShell.
# It runs as a separate process, so Stop Voice can kill it instantly
# and it never gets stuck like pyttsx3 does after being interrupted.
PS_SCRIPT = "; ".join([
    "Add-Type -AssemblyName System.Speech",
    "$s = New-Object System.Speech.Synthesis.SpeechSynthesizer",
    "$l = $env:BPS_LANG",
    "$v = $s.GetInstalledVoices() | Where-Object { $_.Enabled -and "
    "$_.VoiceInfo.Culture.TwoLetterISOLanguageName -eq $l } | "
    "Select-Object -First 1",
    "if ($v) { $s.SelectVoice($v.VoiceInfo.Name) } "
    "elseif ($l -ne 'en') { exit 3 }",
    "$s.Speak($env:BPS_TEXT)",
])

# Offline voice for other systems: pyttsx3 in its own process.
PY_SCRIPT = (
    "import os, sys, pyttsx3\n"
    "words = {'en': ['english', 'en-', 'en_'], 'ta': ['tamil', 'ta-', 'ta_'],"
    " 'hi': ['hindi', 'hi-', 'hi_']}\n"
    "lang = os.environ['BPS_LANG']\n"
    "e = pyttsx3.init()\n"
    "pick = None\n"
    "for v in e.getProperty('voices'):\n"
    "    blob = (str(v.name) + ' ' + str(v.id) + ' ' + "
    "' '.join(str(x) for x in (v.languages or []))).lower()\n"
    "    if any(w in blob for w in words.get(lang, [])):\n"
    "        pick = v\n"
    "        break\n"
    "if pick is None and lang != 'en':\n"
    "    sys.exit(3)\n"
    "if pick is not None:\n"
    "    e.setProperty('voice', pick.id)\n"
    "e.setProperty('rate', 165)\n"
    "e.say(os.environ['BPS_TEXT'])\n"
    "e.runAndWait()\n"
)


class VoiceEngine:
    def __init__(self):
        self.enabled = True
        self.gen = 0
        self.cache_dir = tempfile.mkdtemp(prefix="bps_voice_")
        self.has_gtts = has_module("gtts")
        self.has_pygame = has_module("pygame")
        self.has_pyttsx3 = has_module("pyttsx3")
        self.proc = None
        self.held = False  # True while a button-requested explanation is speaking
        self.on_caption = None
        self.on_status = None
        self.on_notice = None
        self.on_error = None

        if self.has_gtts and (self.has_pygame or IS_WIN):
            self.status_key = "v_status_gtts"
        elif IS_WIN or self.has_pyttsx3:
            self.status_key = "v_status_sys"
        else:
            self.status_key = "v_status_none"

    # ---- public ----
    def speak(self, text, force=False, hold=False):
        """hold=True: a button asked for this explanation, so mouse-hover
        events (tooltips, table, graph) may not interrupt it. Only Stop Voice,
        Voice OFF or another button can."""
        if not (self.enabled or force):
            return
        if not text:
            return
        self.stop()
        self.held = hold
        self.gen += 1
        gen = self.gen
        if self.on_caption:
            self.on_caption(text)
        threading.Thread(
            target=self._run_guard,
            args=(text, lang, gen),
            daemon=True
        ).start()

    def hover_speak(self, text):
        if self.held:
            return
        self.speak(text)

    def hover_stop(self):
        if self.held:
            return
        self.stop()

    def _run_guard(self, text, language, gen):
        try:
            self._run(text, language, gen)
        finally:
            if gen == self.gen:
                self.held = False

    def stop(self):
        self.held = False
        self.gen += 1
        try:
            if self.has_pygame:
                import pygame
                if pygame.mixer.get_init():
                    pygame.mixer.music.stop()
        except Exception:
            pass
        self._kill_proc()

    def cleanup(self):
        self.stop()
        shutil.rmtree(self.cache_dir, ignore_errors=True)

    # ---- internals ----
    def _kill_proc(self):
        proc = self.proc
        self.proc = None
        if proc is not None and proc.poll() is None:
            try:
                proc.kill()
            except Exception:
                pass

    def _set_status(self, key):
        self.status_key = key
        if self.on_status:
            self.on_status()

    def _notice(self, key, **kw):
        if self.on_notice:
            self.on_notice(key, kw)

    def _error(self, message):
        if self.on_error:
            self.on_error(message)

    def _run(self, text, language, gen):
        errors = []

        # 1) Online voice (gTTS) - best Tamil / Hindi quality.
        #    The text is cut into sentences, all are fetched in parallel and
        #    the first one starts playing as soon as it is ready.
        if self.has_gtts and (self.has_pygame or IS_WIN):
            chunks = split_for_speech(text)
            progress = [0]
            try:
                self._speak_online(chunks, language, gen, progress)
                if gen == self.gen:
                    self._set_status("v_status_gtts")
                return
            except Exception as error:
                errors.append("gTTS: %s" % error)
                text = " ".join(chunks[progress[0]:])  # speak only the rest
                if not text:
                    return

        if gen != self.gen:
            return

        # 2) Offline voice installed on this computer
        try:
            self._offline(text, language, gen)
            self._set_status("v_status_sys")
            return
        except NoVoice:
            self._notice("v_nolang", lang=LANG_NAMES.get(language, language))
            return
        except Exception as error:
            errors.append("offline: %s" % error)

        if gen != self.gen:
            return

        self._set_status("v_status_none")
        self._error(" | ".join(errors) or "no voice engine found")

    def _speak_online(self, chunks, language, gen, progress):
        from concurrent.futures import ThreadPoolExecutor

        pool = ThreadPoolExecutor(max_workers=3)
        try:
            futures = [
                pool.submit(self._gtts_file, chunk, language)
                for chunk in chunks
            ]

            for index, future in enumerate(futures):
                while not future.done():
                    if gen != self.gen:
                        return
                    time.sleep(0.03)

                path = future.result()  # raises if gTTS failed

                if gen != self.gen:
                    return

                self._play_file(path, gen, "%d_%d" % (gen, index))
                progress[0] = index + 1
        finally:
            pool.shutdown(wait=False, cancel_futures=True)

    def _gtts_file(self, text, language):
        from gtts import gTTS
        name = hashlib.md5(
            (language + "|" + text).encode("utf-8")
        ).hexdigest() + ".mp3"
        path = os.path.join(self.cache_dir, name)
        if not os.path.exists(path):
            last = None
            for attempt in range(3):
                try:
                    gTTS(text=text, lang=language).save(path)
                    last = None
                    break
                except Exception as error:
                    last = error
                    time.sleep(0.5 * (attempt + 1))
            if last is not None:
                raise last
        return path

    def _play_file(self, path, gen, alias_id=None):
        if self.has_pygame:
            os.environ.setdefault("PYGAME_HIDE_SUPPORT_PROMPT", "1")
            import pygame
            if not pygame.mixer.get_init():
                pygame.mixer.init()
            pygame.mixer.music.load(path)
            pygame.mixer.music.play()
            while pygame.mixer.music.get_busy() and gen == self.gen:
                time.sleep(0.05)
            if gen != self.gen:
                pygame.mixer.music.stop()
            return

        if IS_WIN:
            import ctypes
            mci = ctypes.windll.winmm.mciSendStringW
            alias = "bps%s" % (alias_id or gen)
            mci('open "%s" type mpegvideo alias %s' % (path, alias), None, 0, 0)
            mci("play %s" % alias, None, 0, 0)
            buf = ctypes.create_unicode_buffer(64)
            while gen == self.gen:
                mci("status %s mode" % alias, buf, 63, 0)
                if buf.value != "playing":
                    break
                time.sleep(0.05)
            mci("close %s" % alias, None, 0, 0)
            return

        raise RuntimeError("No audio player")

    def _offline(self, text, language, gen):
        import subprocess

        env = dict(os.environ)
        env["BPS_TEXT"] = text
        env["BPS_LANG"] = language
        flags = 0

        if IS_WIN:
            flags = 0x08000000  # no console window flash
            cmd = [
                "powershell", "-NoProfile", "-NonInteractive",
                "-ExecutionPolicy", "Bypass", "-Command", PS_SCRIPT
            ]
        elif self.has_pyttsx3:
            cmd = [sys.executable, "-c", PY_SCRIPT]
        else:
            raise RuntimeError("no offline voice available")

        proc = subprocess.Popen(
            cmd,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=flags
        )
        self.proc = proc

        _, err = proc.communicate()

        if gen != self.gen:
            return  # stopped on purpose

        if proc.returncode == 3:
            raise NoVoice()

        if proc.returncode != 0:
            detail = err.decode("utf-8", "replace").strip()[:200]
            raise RuntimeError(detail or "exit code %s" % proc.returncode)


voice = VoiceEngine()


# =====================================================================
#                        TOOLTIP (shows + speaks)
# =====================================================================

class ToolTip:
    def __init__(self, widget, key):
        self.widget = widget
        self.key = key
        self.tip = None
        self.job = None

        self.clicked = False

        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.leave, add="+")
        widget.bind("<ButtonPress>", self.press, add="+")

    def text(self):
        return tr(self.key)

    def pointer_inside(self):
        try:
            px, py = self.widget.winfo_pointerxy()
            x0 = self.widget.winfo_rootx()
            y0 = self.widget.winfo_rooty()
            return (
                x0 <= px < x0 + self.widget.winfo_width()
                and y0 <= py < y0 + self.widget.winfo_height()
            )
        except Exception:
            return False

    def press(self, event=None):
        # after a click the tooltip must stay quiet until the mouse leaves
        self.clicked = True
        self.hide()

    def leave(self, event=None):
        # a dialog or window opening can send a fake Leave while the mouse
        # is still on the widget - ignore it
        if self.pointer_inside():
            return
        self.clicked = False
        self.hide()

    def show(self, event=None):
        if self.tip or self.clicked:
            return

        message = self.text()

        x = self.widget.winfo_rootx() + 5
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5

        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")

        tk.Label(
            self.tip,
            text=message,
            bg="#FFF4CC",
            fg="#333333",
            font=(FONT, 9),
            padx=9,
            pady=6,
            wraplength=320,
            justify="left"
        ).pack()

        if voice.enabled:
            self.job = self.widget.after(350, self.say)

    def say(self):
        self.job = None
        voice.hover_speak(self.text())

    def hide(self, event=None):
        if self.job:
            try:
                self.widget.after_cancel(self.job)
            except Exception:
                pass
            self.job = None
        if self.tip:
            self.tip.destroy()
            self.tip = None
            if voice.enabled:
                voice.hover_stop()


# =====================================================================
#                  TEXT REGISTRY (for language switching)
# =====================================================================

i18n_items = []


def reg(widget, key):
    i18n_items.append((widget, key))
    widget.config(text=tr(key))
    return widget


def refresh_texts():
    for widget, key in i18n_items:
        try:
            widget.config(text=tr(key))
        except Exception:
            pass


# =====================================================================
#                              THEME
# =====================================================================

def _theme_walk(parent, bg, fg):
    for child in parent.winfo_children():
        if child is header:
            continue
        if isinstance(child, tk.Frame):
            child.configure(bg=bg)
            _theme_walk(child, bg, fg)
        elif isinstance(child, tk.Label):
            child.configure(bg=bg, fg=fg)


def apply_theme():
    bg = "#202B3C" if dark_mode else BG
    fg = "white" if dark_mode else TEXT
    panel = "#29384D" if dark_mode else "white"

    root.configure(bg=bg)

    _theme_walk(root, bg, fg)

    input_box.configure(bg=panel, fg=fg, insertbackground=fg)

    explanation_label.configure(bg=panel, fg=fg)
    caption_label.configure(
        bg="#33475F" if dark_mode else "#FFF9DB",
        fg="#FFE9A8" if dark_mode else "#5A4A00"
    )

    result_label.configure(bg=bg, fg=fg)

    best_label.configure(
        bg=panel,
        fg="#75D69A" if dark_mode else "#16803C"
    )

    style.configure(
        "Treeview",
        background=panel,
        fieldbackground=panel,
        foreground=fg
    )

    style.configure(
        "Treeview.Heading",
        background="#40516A" if dark_mode else "#DCE8F5",
        foreground=fg
    )

    apply_graph_theme()


def toggle_theme():
    global dark_mode

    dark_mode = not dark_mode
    update_dynamic_buttons()
    apply_theme()


def update_dynamic_buttons():
    theme_button.config(
        text=tr("btn_light") if dark_mode else tr("btn_dark")
    )
    voice_button.config(
        text=tr("btn_voice_on") if voice.enabled else tr("btn_voice_off"),
        bg="#16806A" if voice.enabled else "#9A3B3B",
        activebackground="#16806A" if voice.enabled else "#9A3B3B"
    )
    voice_status_label.config(text=tr(voice.status_key))


# =====================================================================
#                           GRAPH THEME
# =====================================================================

def apply_graph_theme():
    panel = "#29384D" if dark_mode else "white"
    fg = "white" if dark_mode else TEXT

    ax.set_facecolor(panel)
    figure.set_facecolor("#202B3C" if dark_mode else BG)

    ax.title.set_color(fg)
    ax.yaxis.label.set_color(fg)
    ax.tick_params(colors=fg)

    for text in ax.texts:
        text.set_color(fg)

    figure.tight_layout()
    canvas.draw()


# =====================================================================
#                              INPUT
# =====================================================================

def outcome_aliases():
    aliases = {
        "T": "TAKEN",
        "NT": "NOT TAKEN",
        "NOT_TAKEN": "NOT TAKEN",
        "NOT-TAKEN": "NOT TAKEN",
        "TAKEN": "TAKEN",
        "NOT TAKEN": "NOT TAKEN",
    }
    for code in T:
        aliases[T[code]["taken"].upper()] = "TAKEN"
        aliases[T[code]["nottaken"].upper()] = "NOT TAKEN"
    return aliases


def parse_input():
    raw = input_box.get("1.0", tk.END).strip()
    raw = raw.replace("\n", ",")

    if not raw:
        messagebox.showwarning(tr("w_input_t"), tr("w_input"))
        return None

    aliases = outcome_aliases()
    outcomes = []

    for item in raw.split(","):
        item = item.strip().upper()

        if item in aliases:
            outcomes.append(aliases[item])
        else:
            messagebox.showerror(
                tr("e_invalid_t"),
                tr("e_invalid", x=item)
            )
            return None

    return outcomes


# =====================================================================
#                           PREDICTION
# =====================================================================

def predict(actual, method, trace=None):
    """Return the predictions. For the 2-bit counter, `trace` (a list)
    receives (counter_before, counter_after) for every step."""
    predictions = []
    previous = "TAKEN"
    counter = 2  # 0,1 = not taken side   2,3 = taken side (starts Weakly Taken)

    for outcome in actual:

        if method == "Always Taken":
            guess = "TAKEN"

        elif method == "Always Not Taken":
            guess = "NOT TAKEN"

        elif method == "2-Bit Saturating Counter":
            guess = "TAKEN" if counter >= 2 else "NOT TAKEN"
            before = counter

            if outcome == "TAKEN":
                counter = min(3, counter + 1)
            else:
                counter = max(0, counter - 1)

            if trace is not None:
                trace.append((before, counter))

        else:
            guess = previous

        predictions.append(guess)
        previous = outcome

    return predictions


def get_best_methods():
    if not accuracy_data:
        return [], 0

    highest = max(accuracy_data.values())

    best = [
        method
        for method, accuracy in accuracy_data.items()
        if abs(accuracy - highest) < 0.000001
    ]

    return best, highest


def reason_text(item):
    return tr(
        item["rkey"],
        prev=oname(item["prev"]),
        s=tr("s%d" % item.get("state", 0)),
        n=tr("s%d" % item.get("after", 0)),
        p=oname(item["prediction"]),
        a=oname(item["actual"])
    )


# =====================================================================
#                           SIMULATION
# =====================================================================

def compute(actual):
    global last_report, last_accuracies, last_details
    global accuracy_data, results

    last_report = []
    last_accuracies = []
    last_details = []
    accuracy_data = {}
    results = {}

    for method in METHODS:

        trace = []
        predictions = predict(actual, method, trace)
        correct = 0
        details = []
        previous = "TAKEN"

        for number, (real, guess) in enumerate(
            zip(actual, predictions), start=1
        ):

            if method == "Always Taken":
                rkey = "r_t"
            elif method == "Always Not Taken":
                rkey = "r_nt"
            elif method == "2-Bit Saturating Counter":
                rkey = "r_2bit"
            elif number == 1:
                rkey = "r_first"
            else:
                rkey = "r_prev"

            status = "Correct" if real == guess else "Wrong"

            if status == "Correct":
                correct += 1

            item = {
                "method": method,
                "number": number,
                "actual": real,
                "prediction": guess,
                "status": status,
                "rkey": rkey,
                "prev": previous
            }

            if trace:
                item["state"], item["after"] = trace[number - 1]

            details.append(item)
            last_details.append(item)

            last_report.append((
                method, number, real, guess, status
            ))

            previous = real

        accuracy = correct / len(actual) * 100
        accuracy_data[method] = accuracy

        results[method] = {
            "details": details,
            "correct": correct,
            "wrong": len(actual) - correct,
            "accuracy": accuracy
        }

    last_accuracies = [accuracy_data[m] for m in METHODS]


def simulate():
    actual = parse_input()

    if actual is None:
        return

    compute(actual)
    render_results()


def render_results():
    """Draw table, labels and graph in the current language."""
    global best_methods_now, row_info

    table.delete(*table.get_children())
    row_info = {}

    if not results:
        result_label.config(text=tr("ph_result"))
        best_label.config(text=tr("ph_best"))
        explanation_label.config(text=tr("ph_expl"))
        update_graph([0] * len(METHODS), [])
        return

    best_methods_now, highest = get_best_methods()
    summary = []
    explanations = []

    for method in METHODS:

        is_best = method in best_methods_now

        iid = table.insert(
            "",
            tk.END,
            values=(
                mname(method) + (tr("best_tag") if is_best else ""),
                "", "", "", ""
            ),
            tags=(("BestHeading",) if is_best else ("Heading",))
        )
        row_info[iid] = ("head", method)

        for item in results[method]["details"]:

            iid = table.insert(
                "",
                tk.END,
                values=(
                    "",
                    item["number"],
                    oname(item["actual"]),
                    oname(item["prediction"]),
                    sname(item["status"])
                ),
                tags=(item["status"],)
            )
            row_info[iid] = ("row", method, item)

        correct = results[method]["correct"]
        wrong = results[method]["wrong"]
        accuracy = results[method]["accuracy"]
        total = correct + wrong

        summary.append(f"{mname(method)}: {correct}/{total} ({accuracy:.1f}%)")
        explanations.append(
            tr("expl_line", m=mname(method), c=correct, w=wrong)
        )

    result_label.config(text=tr("res_head") + "\n" + "\n".join(summary))

    if len(best_methods_now) == 1:
        best_text = tr("best_one", m=mname(best_methods_now[0]), a=highest)
    else:
        best_text = tr(
            "best_tie",
            m=", ".join(mname(m) for m in best_methods_now),
            a=highest
        )

    best_label.config(text=best_text)
    explanation_label.config(
        text=tr("expl_head") + "\n" + "\n".join(explanations)
    )

    update_graph(last_accuracies, best_methods_now)
    table.yview_moveto(0)


# =====================================================================
#                               GRAPH
# =====================================================================

def update_graph(accuracies, best_methods=None):
    global bar_objects

    ax.clear()

    if best_methods is None:
        best_methods = []

    bars = ax.bar(
        [mname(m) for m in METHODS],
        accuracies,
        color=COLORS
    )
    bar_objects = list(bars)

    for method, bar in zip(METHODS, bars):
        if method in best_methods:
            bar.set_edgecolor("#16803C")
            bar.set_linewidth(3)

    ax.set_title(tr("graph_t"), fontsize=11, fontweight="bold")
    ax.set_ylabel(tr("graph_y"))
    ax.set_ylim(0, 110)

    if results:
        ax.bar_label(bars, fmt="%.1f%%", padding=3)

    ax.tick_params(axis="x", labelrotation=18, labelsize=7)

    apply_graph_theme()


graph_hover = {"last": None, "job": None}


def graph_motion(event):
    if event.inaxes != ax:
        graph_leave(None)
        return

    key = "graph"
    for index, bar in enumerate(bar_objects):
        if bar.contains(event)[0]:
            key = index
            break

    if key == graph_hover["last"]:
        return

    graph_hover["last"] = key

    if graph_hover["job"]:
        root.after_cancel(graph_hover["job"])
        graph_hover["job"] = None

    if not voice.enabled:
        return

    graph_hover["job"] = root.after(250, lambda: graph_say(key))


def graph_say(key):
    graph_hover["job"] = None

    if key == "graph" or not results:
        voice.hover_speak(tr("tip_graph"))
        return

    method = METHODS[key]
    text = tr(
        "head_voice",
        m=mname(method),
        d=mdesc(method),
        c=results[method]["correct"],
        w=results[method]["wrong"],
        a=results[method]["accuracy"]
    )
    if method in best_methods_now:
        text += tr("best_voice")
    voice.hover_speak(text)


def graph_leave(event):
    graph_hover["last"] = None
    if graph_hover["job"]:
        root.after_cancel(graph_hover["job"])
        graph_hover["job"] = None


# =====================================================================
#                          TABLE HOVER VOICE
# =====================================================================

table_hover = {"last": None, "job": None}


def row_voice_text(iid):
    info = row_info.get(iid)
    if not info:
        return None

    if info[0] == "head":
        method = info[1]
        text = tr(
            "head_voice",
            m=mname(method),
            d=mdesc(method),
            c=results[method]["correct"],
            w=results[method]["wrong"],
            a=results[method]["accuracy"]
        )
        if method in best_methods_now:
            text += tr("best_voice")
        return text

    item = info[2]
    return tr(
        "row_voice",
        n=item["number"],
        a=oname(item["actual"]),
        p=oname(item["prediction"]),
        reason=reason_text(item),
        s=sname(item["status"])
    )


def table_motion(event):
    iid = table.identify_row(event.y)

    if iid == table_hover["last"]:
        return

    table_hover["last"] = iid

    if table_hover["job"]:
        root.after_cancel(table_hover["job"])
        table_hover["job"] = None

    if not voice.enabled:
        return

    if not iid:
        return

    table_hover["job"] = root.after(300, lambda: table_say(iid))


def table_say(iid):
    table_hover["job"] = None

    if not results:
        voice.hover_speak(tr("tip_table"))
        return

    text = row_voice_text(iid)
    if text:
        voice.hover_speak(text)


def table_enter(event):
    if not results and voice.enabled:
        table_hover["job"] = root.after(
            350, lambda: voice.hover_speak(tr("tip_table"))
        )


def table_leave(event):
    table_hover["last"] = None
    if table_hover["job"]:
        root.after_cancel(table_hover["job"])
        table_hover["job"] = None
    if voice.enabled:
        voice.hover_stop()


# =====================================================================
#                         VOICE BUTTON ACTIONS
# =====================================================================

def toggle_voice():
    voice.enabled = not voice.enabled

    if not voice.enabled:
        voice.stop()
    else:
        voice.speak(tr("voice_enabled"), hold=True)

    update_dynamic_buttons()


def stop_voice():
    voice.stop()


def explain_methods():
    parts = [tr("explain_intro")]
    for method in METHODS:
        parts.append(mname(method) + ". " + mdesc(method))
    voice.speak("  ".join(parts), force=True, hold=True)


def read_results():
    if not results:
        voice.speak(tr("no_results_voice"), force=True, hold=True)
        return

    parts = [tr("read_intro")]

    for method in METHODS:
        parts.append(tr(
            "head_voice",
            m=mname(method),
            d="",
            c=results[method]["correct"],
            w=results[method]["wrong"],
            a=results[method]["accuracy"]
        ).replace("  ", " "))

    highest = max(accuracy_data.values())
    if len(best_methods_now) == 1:
        parts.append(tr("best_one", m=mname(best_methods_now[0]), a=highest))
    else:
        parts.append(tr(
            "best_tie",
            m=", ".join(mname(m) for m in best_methods_now),
            a=highest
        ))

    voice.speak("  ".join(parts), force=True, hold=True)


def show_caption(text):
    root.after(0, lambda: caption_label.config(text=text))


def show_notice(key, kw):
    root.after(0, lambda: set_voice_message(tr(key, **kw)))


info_win = {"win": None, "status": None}

SPEECH_WORDS = {
    "en": (" times ", " equals ", " divided by "),
    "ta": (" பெருக்கல் ", " சமம் ", " வகுத்தல் "),
    "hi": (" गुणा ", " बराबर ", " भाग "),
}


def speech_clean(text):
    """Make a multi-line text pleasant to hear: pauses at line ends and
    spoken words for symbols such as = and /."""
    lines = [line.strip() for line in text.split("\n") if line.strip()]
    parts = [
        line if line[-1] in ".!?:\u0964" else line + "."
        for line in lines
    ]
    joined = " ".join(parts)
    times, equals, by = SPEECH_WORDS.get(lang, SPEECH_WORDS["en"])
    joined = joined.replace(" x ", times).replace("=", equals)
    joined = joined.replace(" / ", by)
    return " ".join(joined.split())


def set_voice_message(text):
    caption_label.config(text=text)
    label = info_win.get("status")
    try:
        if label is not None and label.winfo_exists():
            label.config(text=text)
    except Exception:
        pass


def show_info_window(title_key, text_key):
    """Non-blocking information window that also reads the text aloud."""
    old = info_win["win"]
    if old is not None:
        try:
            old.destroy()
        except Exception:
            pass

    text = tr(text_key)
    spoken = speech_clean(text)

    bg = "#202B3C" if dark_mode else BG
    fg = "white" if dark_mode else TEXT
    panel = "#29384D" if dark_mode else "white"

    win = tk.Toplevel(root)
    win.title(tr(title_key))
    win.geometry("640x560")
    win.configure(bg=bg)
    win.transient(root)

    tk.Label(
        win, text=tr(title_key), font=(FONT, 14, "bold"), bg=bg, fg=fg
    ).pack(pady=(12, 4))

    frame = tk.Frame(win, bg=bg)
    frame.pack(fill="both", expand=True, padx=12, pady=6)

    scroll = ttk.Scrollbar(frame, orient="vertical")
    box = tk.Text(
        frame, wrap="word", font=(FONT, 11), bg=panel, fg=fg,
        padx=12, pady=10, yscrollcommand=scroll.set
    )
    scroll.config(command=box.yview)
    scroll.pack(side="right", fill="y")
    box.pack(side="left", fill="both", expand=True)
    box.insert("1.0", text)
    box.config(state="disabled")

    status = tk.Label(
        win, text="", font=(FONT, 9, "italic"), bg=bg, fg="#D98A00",
        wraplength=600, justify="left"
    )
    status.pack(fill="x", padx=14)

    def close():
        voice.stop()
        info_win["win"] = None
        info_win["status"] = None
        win.destroy()

    bar = tk.Frame(win, bg=bg)
    bar.pack(pady=8)

    for key, color, command in (
        ("btn_read_aloud", "#16806A", lambda: voice.speak(spoken, force=True, hold=True)),
        ("btn_stop", "#9A3B3B", stop_voice),
        ("btn_close", "#555555", close),
    ):
        tk.Button(
            bar, text=tr(key), command=command, bg=color, fg="white",
            activebackground=color, activeforeground="white",
            font=(FONT, 9, "bold"), padx=12, pady=5, cursor="hand2",
            relief="flat"
        ).pack(side="left", padx=4)

    win.protocol("WM_DELETE_WINDOW", close)

    info_win["win"] = win
    info_win["status"] = status

    # speak after the window is on screen (no blocking dialog involved)
    win.after(250, lambda: voice.speak(spoken, hold=True))


def info_box(title_key, text_key):
    show_info_window(title_key, text_key)


# =====================================================================
#                         LANGUAGE SWITCH
# =====================================================================

def change_language(event=None):
    global lang

    lang = LANGS.get(lang_box.get(), "en")

    root.title(tr("title"))
    refresh_texts()
    update_dynamic_buttons()

    for col, key in zip(
        columns,
        ("col_method", "col_no", "col_actual", "col_pred", "col_status")
    ):
        table.heading(col, text=tr(key))

    welcome_text_label.config(text=tr("welcome_text"))

    if results:
        render_results()
    else:
        render_results()

    caption_label.config(text=tr("ph_caption"))

    voice.speak(tr("lang_changed"), force=True, hold=True)


# =====================================================================
#                      STEP-BY-STEP EXPLANATION
# =====================================================================

def show_steps():

    if not last_details:
        messagebox.showwarning(tr("w_nores_t"), tr("w_nores"))
        return

    window = tk.Toplevel(root)
    window.title(tr("steps_title"))
    window.geometry("720x620")
    window.configure(bg="#202B3C" if dark_mode else BG)

    fg = "white" if dark_mode else TEXT
    panel = "#29384D" if dark_mode else "white"

    tk.Label(
        window,
        text=tr("steps_head"),
        font=(FONT, 14, "bold"),
        bg="#202B3C" if dark_mode else BG,
        fg=fg
    ).pack(pady=10)

    text_frame = tk.Frame(window)
    text_frame.pack(fill="both", expand=True, padx=12, pady=8)

    scroll = ttk.Scrollbar(text_frame, orient="vertical")

    details_box = tk.Text(
        text_frame,
        wrap="word",
        font=(FONT, 10),
        bg=panel,
        fg=fg,
        padx=10,
        pady=10,
        yscrollcommand=scroll.set
    )

    scroll.config(command=details_box.yview)
    scroll.pack(side="right", fill="y")
    details_box.pack(side="left", fill="both", expand=True)

    spoken = []

    for method in METHODS:

        details_box.insert(tk.END, f"\n{mname(method)}\n", "heading")
        spoken.append(mname(method) + ".")

        for item in last_details:

            if item["method"] != method:
                continue

            details_box.insert(
                tk.END,
                f"\n{tr('step_n', n=item['number'])}\n"
                f"{tr('lbl_actual')}: {oname(item['actual'])}\n"
                f"{tr('lbl_pred')}: {oname(item['prediction'])}\n"
                f"{tr('lbl_reason')}: {reason_text(item)}\n"
                f"{tr('lbl_result')}: {sname(item['status'])}\n"
            )

            spoken.append(tr(
                "row_voice",
                n=item["number"],
                a=oname(item["actual"]),
                p=oname(item["prediction"]),
                reason=reason_text(item),
                s=sname(item["status"])
            ))

    details_box.tag_configure(
        "heading",
        font=(FONT, 11, "bold"),
        foreground="#55A9E8"
    )

    details_box.config(state="disabled")

    bar = tk.Frame(window, bg="#202B3C" if dark_mode else BG)
    bar.pack(pady=6)

    tk.Button(
        bar,
        text=tr("btn_read_aloud"),
        command=lambda: voice.speak("  ".join(spoken), force=True, hold=True),
        bg="#16806A", fg="white", relief="flat",
        font=(FONT, 9, "bold"), padx=10, pady=5, cursor="hand2"
    ).pack(side="left", padx=4)

    tk.Button(
        bar,
        text=tr("btn_stop"),
        command=stop_voice,
        bg="#9A3B3B", fg="white", relief="flat",
        font=(FONT, 9, "bold"), padx=10, pady=5, cursor="hand2"
    ).pack(side="left", padx=4)

    window.protocol(
        "WM_DELETE_WINDOW",
        lambda: (voice.stop(), window.destroy())
    )


# =====================================================================
#                          INFO WINDOWS
# =====================================================================

def show_formula():
    info_box("formula_t", "formula_m")


def about():
    info_box("about_t", "about_m")


def how_to_use():
    info_box("how_t", "how_m")


# =====================================================================
#                              RESET
# =====================================================================

def reset():
    global last_report, last_accuracies, last_details
    global accuracy_data, results, best_methods_now

    input_box.delete("1.0", tk.END)
    input_box.insert(tk.END, DEFAULT_INPUT)

    last_report = []
    last_accuracies = []
    last_details = []
    accuracy_data = {}
    results = {}
    best_methods_now = []

    render_results()


# =====================================================================
#                          SAMPLE INPUT
# =====================================================================

def sample_input():
    input_box.delete("1.0", tk.END)

    input_box.insert(
        tk.END,
        "TAKEN, TAKEN, NOT TAKEN, NOT TAKEN"
    )

    show_info_window("sample_t", "sample_m")


# =====================================================================
#                           CSV REPORT
# =====================================================================

def download_csv():

    if not last_report:
        messagebox.showwarning(tr("w_nores_t"), tr("w_nores_dl"))
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".csv",
        filetypes=[("CSV files", "*.csv")],
        title=tr("save_csv_t")
    )

    if not path:
        return

    try:
        with open(path, "w", newline="", encoding="utf-8-sig") as file:

            writer = csv.writer(file)

            writer.writerow([
                "Method",
                "Prediction Number",
                "Actual Outcome",
                "Predicted Outcome",
                "Status",
                "Accuracy (%)"
            ])

            for row in last_report:
                accuracy = accuracy_data[row[0]]
                writer.writerow([
                    row[0], row[1], row[2], row[3], row[4],
                    f"{accuracy:.1f}"
                ])

        messagebox.showinfo(tr("csv_saved_t"), tr("csv_saved"))

    except Exception as error:
        messagebox.showerror(tr("err_t"), str(error))


# =====================================================================
#                            PDF REPORT
# =====================================================================

def download_report():

    if not last_report:
        messagebox.showwarning(tr("w_nores_t"), tr("w_nores_dl"))
        return

    path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        title=tr("save_pdf_t")
    )

    if not path:
        return

    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas as pdf_canvas

        pdf = pdf_canvas.Canvas(path, pagesize=A4)

        width, height = A4
        y = height - 55

        pdf.setFont("Helvetica-Bold", 16)
        pdf.drawString(45, y, "BRANCH PREDICTION SIMULATOR")
        y -= 30

        pdf.setFont("Helvetica", 10)
        pdf.drawString(45, y, "Simulation Results")
        y -= 25

        for method in METHODS:

            if y < 70:
                pdf.showPage()
                y = height - 55

            pdf.setFont("Helvetica-Bold", 11)
            pdf.drawString(45, y, method)
            y -= 18

            for row in last_report:

                if row[0] == method:

                    if y < 50:
                        pdf.showPage()
                        y = height - 55

                    line = (
                        f"No: {row[1]} | "
                        f"Actual: {row[2]} | "
                        f"Prediction: {row[3]} | "
                        f"{row[4]}"
                    )

                    pdf.setFont("Helvetica", 9)
                    pdf.drawString(55, y, line)
                    y -= 15

            y -= 8

        if y < 70:
            pdf.showPage()
            y = height - 55

        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawString(45, y, "ACCURACY SUMMARY")
        y -= 20

        for method in METHODS:
            pdf.setFont("Helvetica", 10)
            pdf.drawString(45, y, f"{method}: {accuracy_data[method]:.1f}%")
            y -= 18

        best_methods, highest = get_best_methods()

        if y < 60:
            pdf.showPage()
            y = height - 55

        pdf.setFont("Helvetica-Bold", 11)

        if len(best_methods) == 1:
            best_text = f"Best Method: {best_methods[0]} ({highest:.1f}%)"
        else:
            best_text = (
                f"Best Methods (Tie): {', '.join(best_methods)} "
                f"({highest:.1f}%)"
            )

        pdf.drawString(45, y, best_text)
        pdf.save()

        messagebox.showinfo(tr("pdf_saved_t"), tr("pdf_saved"))

    except ImportError:
        messagebox.showerror(tr("lib_t"), tr("lib_m"))

    except Exception as error:
        messagebox.showerror(tr("err_t"), str(error))


# =====================================================================
#                         BUTTON CREATION
# =====================================================================

def make_button(parent, key, command, color, tip_key):

    button = tk.Button(
        parent,
        command=command,
        bg=color,
        fg="white",
        activebackground=color,
        activeforeground="white",
        font=(FONT, 9, "bold"),
        padx=10,
        pady=6,
        cursor="hand2",
        relief="flat"
    )

    reg(button, key)
    button.pack(side="left", padx=4)

    ToolTip(button, tip_key)

    return button


def make_label(parent, key, **kw):
    label = tk.Label(parent, **kw)
    reg(label, key)
    return label


# =====================================================================
#                           MAIN WINDOW
# =====================================================================

root = tk.Tk()

_families = set(tkfont.families())
for _name in ("Nirmala UI", "Noto Sans Tamil", "Noto Sans Devanagari",
              "Latha", "Mangal"):
    if _name in _families:
        FONT = _name
        break

root.title(tr("title"))
root.geometry("1250x900")
root.minsize(1000, 740)
root.configure(bg=BG)


# ---------------- HEADER ----------------

header = tk.Frame(root, bg=BLUE)
header.pack(fill="x", padx=12, pady=(10, 5))

logo = tk.Canvas(
    header,
    width=65,
    height=65,
    bg=BLUE,
    highlightthickness=0
)

logo.pack(side="left", padx=12, pady=5)

logo.create_rectangle(17, 17, 48, 48, outline="white", width=3)
logo.create_rectangle(23, 23, 42, 42, outline="#8ED0FF", width=2)

for pos in (22, 32, 42):
    logo.create_line(pos, 9, pos, 17, fill="white", width=2)
    logo.create_line(pos, 48, pos, 56, fill="white", width=2)
    logo.create_line(9, pos, 17, pos, fill="white", width=2)
    logo.create_line(48, pos, 56, pos, fill="white", width=2)

heading = tk.Frame(header, bg=BLUE)
heading.pack(side="left", fill="y", pady=10)

title_label = tk.Label(
    heading, font=(FONT, 19, "bold"), bg=BLUE, fg="white"
)
reg(title_label, "title")
title_label.pack(anchor="w")

tagline_label = tk.Label(
    heading, font=(FONT, 10), bg=BLUE, fg="#D8E9FA"
)
reg(tagline_label, "tagline")
tagline_label.pack(anchor="w")

for _w in (header, logo, heading, title_label, tagline_label):
    ToolTip(_w, "tip_header")


# ---------------- WELCOME ----------------

welcome = tk.Frame(root, bg=BG)
welcome.pack(fill="x", padx=18, pady=5)

make_label(
    welcome, "welcome_title",
    font=(FONT, 14, "bold"), bg=BG, fg=TEXT
).pack(anchor="w")

welcome_text_label = tk.Label(
    welcome,
    text=tr("welcome_text"),
    font=(FONT, 9),
    bg=BG,
    fg=TEXT,
    wraplength=1100,
    justify="left"
)
welcome_text_label.pack(anchor="w")


# ---------------- INFORMATION BUTTONS ----------------

info_frame = tk.Frame(root, bg=BG)
info_frame.pack(pady=4)

make_button(info_frame, "btn_about", about, "#16806A", "tip_about")
make_button(info_frame, "btn_how", how_to_use, "#2878B5", "tip_how")
make_button(info_frame, "btn_formula", show_formula, "#8064B4", "tip_formula")

theme_button = tk.Button(
    info_frame, command=toggle_theme, bg="#555555", fg="white",
    activebackground="#555555", activeforeground="white",
    font=(FONT, 9, "bold"), padx=10, pady=6, cursor="hand2", relief="flat"
)
theme_button.pack(side="left", padx=4)
ToolTip(theme_button, "tip_theme")


# ---------------- LANGUAGE + VOICE BAR ----------------

voice_frame = tk.Frame(root, bg=BG)
voice_frame.pack(pady=4)

make_label(
    voice_frame, "lang_label",
    font=(FONT, 10, "bold"), bg=BG, fg=TEXT
).pack(side="left", padx=(0, 4))

lang_box = ttk.Combobox(
    voice_frame,
    values=list(LANGS.keys()),
    state="readonly",
    width=10,
    font=(FONT, 10)
)
lang_box.set("English")
lang_box.pack(side="left", padx=(0, 10))
lang_box.bind("<<ComboboxSelected>>", change_language)
ToolTip(lang_box, "tip_lang")

voice_button = tk.Button(
    voice_frame, command=toggle_voice, fg="white",
    activeforeground="white", font=(FONT, 9, "bold"),
    padx=10, pady=6, cursor="hand2", relief="flat"
)
voice_button.pack(side="left", padx=4)
ToolTip(voice_button, "tip_voice")

make_button(voice_frame, "btn_explain", explain_methods, "#2463A6", "tip_explain")
make_button(voice_frame, "btn_read", read_results, "#8064B4", "tip_read")
make_button(voice_frame, "btn_stop", stop_voice, "#9A3B3B", "tip_stop")

voice_status_label = tk.Label(
    voice_frame, font=(FONT, 8, "italic"), bg=BG, fg="#777777"
)
voice_status_label.pack(side="left", padx=10)


# ---------------- INPUT AREA ----------------

input_frame = tk.Frame(root, bg=BG)
input_frame.pack(fill="x", padx=18, pady=7)

make_label(
    input_frame, "input_label",
    font=(FONT, 10, "bold"), bg=BG, fg=TEXT
).pack(anchor="w")

input_box = tk.Text(
    input_frame,
    height=2,
    font=(FONT, 10),
    wrap="word"
)

input_box.pack(fill="x", pady=5)
input_box.insert(tk.END, DEFAULT_INPUT)

ToolTip(input_box, "tip_input")


# ---------------- SIMULATION BUTTONS ----------------

button_frame = tk.Frame(root, bg=BG)
button_frame.pack(pady=5)

make_button(button_frame, "btn_run", simulate, "#2463A6", "tip_run")
make_button(button_frame, "btn_sample", sample_input, "#8064B4", "tip_sample")
make_button(button_frame, "btn_reset", reset, "#555555", "tip_reset")
make_button(button_frame, "btn_steps", show_steps, "#B07827", "tip_steps")
make_button(button_frame, "btn_pdf", download_report, "#16806A", "tip_pdf")
make_button(button_frame, "btn_csv", download_csv, "#16806A", "tip_csv")


# ---------------- MAIN CONTENT ----------------

content = tk.Frame(root, bg=BG)

content.pack(fill="both", expand=True, padx=14, pady=8)

content.columnconfigure(0, weight=6)
content.columnconfigure(1, weight=5)
content.rowconfigure(0, weight=1)


# ---------------- RESULTS TABLE ----------------

left_panel = tk.Frame(content, bg=BG)
left_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 8))

make_label(
    left_panel, "table_title",
    font=(FONT, 12, "bold"), bg=BG, fg=TEXT
).pack(anchor="w", pady=5)

table_frame = tk.Frame(left_panel)
table_frame.pack(fill="both", expand=True)

columns = ("Method", "No", "Actual", "Prediction", "Status")

style = ttk.Style()
style.theme_use("clam")

style.configure("Treeview", rowheight=25, font=(FONT, 9))
style.configure("Treeview.Heading", font=(FONT, 9, "bold"))

table = ttk.Treeview(table_frame, columns=columns, show="headings")

for col, key in zip(
    columns,
    ("col_method", "col_no", "col_actual", "col_pred", "col_status")
):
    table.heading(col, text=tr(key))

table.column("Method", width=170, anchor="center")
table.column("No", width=40, anchor="center")
table.column("Actual", width=105, anchor="center")
table.column("Prediction", width=105, anchor="center")
table.column("Status", width=75, anchor="center")

table_scroll = ttk.Scrollbar(
    table_frame, orient="vertical", command=table.yview
)

table.configure(yscrollcommand=table_scroll.set)

table.pack(side="left", fill="both", expand=True)
table_scroll.pack(side="right", fill="y")

table.tag_configure(
    "Heading", background="#DCE8F5", foreground="#183153",
    font=(FONT, 9, "bold")
)
table.tag_configure(
    "BestHeading", background="#BDE8CC", foreground="#125C30",
    font=(FONT, 9, "bold")
)
table.tag_configure("Correct", foreground="#16803C")
table.tag_configure("Wrong", foreground="#D62828")

table.bind("<Motion>", table_motion)
table.bind("<Enter>", table_enter)
table.bind("<Leave>", table_leave)


# ---------------- GRAPH PANEL ----------------

right_panel = tk.Frame(content, bg=BG)
right_panel.grid(row=0, column=1, sticky="nsew", padx=(8, 0))

make_label(
    right_panel, "graph_lbl",
    font=(FONT, 12, "bold"), bg=BG, fg=TEXT
).pack(anchor="w", pady=5)

figure = Figure(figsize=(5.2, 4.5), dpi=100)

ax = figure.add_subplot(111)

canvas = FigureCanvasTkAgg(figure, master=right_panel)

canvas.get_tk_widget().pack(fill="both", expand=True)

canvas.mpl_connect("motion_notify_event", graph_motion)
canvas.mpl_connect("axes_leave_event", graph_leave)
canvas.mpl_connect("figure_leave_event", graph_leave)
canvas.get_tk_widget().bind(
    "<Leave>", lambda e: (graph_leave(e), voice.hover_stop() if voice.enabled else None)
)


# ---------------- ACCURACY SUMMARY ----------------

result_label = tk.Label(
    right_panel,
    font=(FONT, 9, "bold"),
    bg=BG,
    fg=TEXT,
    justify="left",
    anchor="w",
    wraplength=450
)

result_label.pack(fill="x", pady=6)
ToolTip(result_label, "tip_result")


# ---------------- BEST METHOD ----------------

best_label = tk.Label(
    right_panel,
    font=(FONT, 10, "bold"),
    bg="white",
    fg="#16803C",
    justify="left",
    anchor="w",
    wraplength=450,
    padx=8,
    pady=7
)

best_label.pack(fill="x", pady=4)
ToolTip(best_label, "tip_best")


# ---------------- EXPLANATION ----------------

explanation_label = tk.Label(
    root,
    font=(FONT, 9),
    bg="#E4ECF5",
    fg=TEXT,
    justify="left",
    anchor="w",
    padx=12,
    pady=8,
    wraplength=1150
)

explanation_label.pack(fill="x", padx=15, pady=(2, 4))
ToolTip(explanation_label, "tip_expl")


# ---------------- VOICE CAPTION ----------------

caption_label = tk.Label(
    root,
    text=tr("ph_caption"),
    font=(FONT, 9, "italic"),
    bg="#FFF9DB",
    fg="#5A4A00",
    justify="left",
    anchor="w",
    padx=12,
    pady=6,
    wraplength=1150
)

caption_label.pack(fill="x", padx=15, pady=(2, 6))
ToolTip(caption_label, "tip_caption")


# ---------------- FOOTER ----------------

tk.Label(
    root,
    text="Simple Learning • Easy Simulation • Clear Results",
    font=(FONT, 8, "italic"),
    bg=BG,
    fg="#777777"
).pack(pady=(0, 8))


# ---------------- WIRE UP VOICE CALLBACKS ----------------

voice.on_caption = show_caption
voice.on_notice = show_notice
voice.on_error = lambda m: root.after(
    0, lambda: set_voice_message("Voice error: " + m)
)
voice.on_status = lambda: root.after(0, update_dynamic_buttons)


def on_close():
    voice.cleanup()
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_close)


# ---------------- START APPLICATION ----------------

update_dynamic_buttons()
render_results()
apply_theme()

root.mainloop()