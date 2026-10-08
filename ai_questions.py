"""
AI Question Generator — Gemini-Powered with Rule-Based Fallback
Generates course-related proctoring questions based on subject context
"""
import random
import json
from functools import lru_cache

try:
    import google.generativeai as genai
except ImportError:
    pass

QUESTION_BANK = {
    'computer': [
        {"q": "What does CPU stand for?", "options": ["Central Processing Unit", "Computer Personal Unit", "Core Processing Unit", "Central Program Unit"], "answer": 0},
        {"q": "Which data structure uses LIFO?", "options": ["Queue", "Stack", "Array", "Tree"], "answer": 1},
        {"q": "What is the time complexity of binary search?", "options": ["O(n)", "O(n²)", "O(log n)", "O(1)"], "answer": 2},
        {"q": "HTML stands for?", "options": ["Hyper Text Markup Language", "High Tech Modern Language", "Hyper Transfer Markup Language", "None"], "answer": 0},
        {"q": "Which language is used for web styling?", "options": ["Java", "Python", "CSS", "C++"], "answer": 2},
        {"q": "What does RAM stand for?", "options": ["Read Access Memory", "Random Access Memory", "Rapid Access Module", "Read All Memory"], "answer": 1},
        {"q": "Which sorting algorithm has best average complexity?", "options": ["Bubble Sort", "Selection Sort", "Quick Sort", "Insertion Sort"], "answer": 2},
        {"q": "What is an IP address?", "options": ["Internet Protocol address", "Internal Port address", "Internet Process address", "None"], "answer": 0},
    ],
    'math': [
        {"q": "What is the derivative of x²?", "options": ["x", "2x", "x²", "2"], "answer": 1},
        {"q": "What is the value of π (pi) approximately?", "options": ["3.14", "2.71", "1.61", "1.41"], "answer": 0},
        {"q": "What is the integral of 1/x?", "options": ["x", "ln(x)", "1/x²", "e^x"], "answer": 1},
        {"q": "Which is a prime number?", "options": ["9", "15", "17", "21"], "answer": 2},
        {"q": "What is log(1)?", "options": ["1", "0", "undefined", "∞"], "answer": 1},
        {"q": "What is 2^10?", "options": ["512", "1024", "256", "2048"], "answer": 1},
    ],
    'physics': [
        {"q": "What is the speed of light?", "options": ["3×10⁸ m/s", "3×10⁶ m/s", "3×10¹⁰ m/s", "3×10⁴ m/s"], "answer": 0},
        {"q": "F = ma is Newton's which law?", "options": ["First", "Second", "Third", "Fourth"], "answer": 1},
        {"q": "Unit of electric current?", "options": ["Volt", "Watt", "Ampere", "Ohm"], "answer": 2},
        {"q": "What is the unit of force?", "options": ["Joule", "Newton", "Pascal", "Watt"], "answer": 1},
        {"q": "Energy = ?", "options": ["mc", "mc²", "m²c", "m/c"], "answer": 1},
    ],
    'chemistry': [
        {"q": "What is the chemical symbol for Gold?", "options": ["Go", "Gd", "Au", "Ag"], "answer": 2},
        {"q": "Water is made of?", "options": ["H₂O₂", "HO", "H₂O", "H₃O"], "answer": 2},
        {"q": "Atomic number of Carbon?", "options": ["4", "6", "8", "12"], "answer": 1},
        {"q": "pH of pure water?", "options": ["5", "6", "7", "8"], "answer": 2},
        {"q": "Which gas do plants absorb?", "options": ["Oxygen", "Nitrogen", "CO₂", "Hydrogen"], "answer": 2},
    ],
    'default': [
        {"q": "Are you paying attention to the lecture?", "options": ["Yes, fully focused", "Mostly yes", "Somewhat", "Not really"], "answer": 0},
        {"q": "What is the main topic being discussed right now?", "options": ["I was listening", "I missed it", "I was distracted", "I don't know"], "answer": 0},
        {"q": "Have you taken notes in the last 5 minutes?", "options": ["Yes", "No", "Partially", "Will do now"], "answer": 0},
        {"q": "Can you summarize the last point made?", "options": ["Yes I can", "Partially", "I missed it", "I was not listening"], "answer": 0},
    ]
}

@lru_cache(maxsize=32)
def get_questions_for_course(course_name, course_code, count=5, api_key=None):
    """Get dynamic AI questions with rule-based fallback"""
    
    # 1. Try Gemini API
    if api_key:
        try:
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')
            prompt = f"""You are an academic AI generating proctoring questions for an active online class.
Course: '{course_name}' (Code: {course_code})
Generate EXACTLY {count} multiple-choice questions covering fundamental concepts of this specific course. Keep questions concise.

Return ONLY valid JSON in this exact structure:
[
  {{"q": "Question text?", "options": ["Option A", "Option B", "Option C", "Option D"], "answer": 0}}
]
The 'answer' should be the integer index (0-3) of the correct option. Do not include markdown codeblocks or any standard conversational text."""

            response = model.generate_content(prompt)
            text = response.text.strip()
            
            # Clean up markdown if the LLM adds it anyway
            if text.startswith("```json"): text = text[7:]
            elif text.startswith("```"): text = text[3:]
            if text.endswith("```"): text = text[:-3]
            
            questions = json.loads(text.strip())
            return [{"id": i, **q} for i, q in enumerate(questions[:count])]
            
        except Exception as e:
            print(f"[WARN] Gemini AI failed for questions, falling back to dict: {str(e)}")

    # 2. Rule-based Fallback
    course_lower = (course_name + course_code).lower()

    if any(w in course_lower for w in ['computer', 'cs', 'it', 'software', 'programming', 'data']):
        bank = QUESTION_BANK['computer']
    elif any(w in course_lower for w in ['math', 'calculus', 'algebra', 'statistics']):
        bank = QUESTION_BANK['math']
    elif any(w in course_lower for w in ['physics', 'phy', 'mechanics']):
        bank = QUESTION_BANK['physics']
    elif any(w in course_lower for w in ['chemistry', 'chem', 'organic']):
        bank = QUESTION_BANK['chemistry']
    else:
        bank = QUESTION_BANK['default']

    # Mix with default questions
    combined = bank + QUESTION_BANK['default']
    selected = random.sample(combined, min(count, len(combined)))

    return [{"id": i, **q} for i, q in enumerate(selected)]


def get_random_intervals(num_questions, class_duration_minutes=60):
    """
    Generate unpredictable random intervals for questions.
    Returns list of seconds after class start when each question fires.
    """
    if num_questions == 0:
        return []

    # Divide class into segments and pick random time within each segment
    segment_size = (class_duration_minutes * 60) // (num_questions + 1)
    intervals = []

    for i in range(num_questions):
        start = segment_size * (i + 1) - segment_size // 2
        end = segment_size * (i + 1) + segment_size // 2
        # Add randomness — students can't predict exact time
        interval = random.randint(max(60, start), min(class_duration_minutes * 60 - 60, end))
        intervals.append(interval)

    intervals.sort()
    return intervals