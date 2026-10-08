"""
AI Features Module
==================
1. Grade Predictor - predicts final grade based on attendance + current grades
2. Dropout Risk Detector - identifies students likely to drop out
3. Personalized Study Recommendations - suggests study plans based on weak areas
"""

# ============================================================================
# 1. GRADE PREDICTOR
# ============================================================================

def predict_final_grade(attendance_pct, current_score, assignments_submitted, total_assignments):
    """
    Predicts final grade based on:
    - Attendance percentage
    - Current score/grade
    - Assignment completion rate

    Returns: predicted grade letter, predicted score, confidence
    """
    if total_assignments == 0:
        assignment_rate = 0
    else:
        assignment_rate = (assignments_submitted / total_assignments) * 100

    # Weighted score calculation
    # Attendance: 20%, Current score: 60%, Assignment rate: 20%
    predicted_score = (
        (attendance_pct * 0.20) +
        (current_score * 0.60) +
        (assignment_rate * 0.20)
    )
    predicted_score = round(predicted_score, 1)

    # Convert score to grade
    if predicted_score >= 90:
        grade = 'A'
        grade_color = '#198754'
    elif predicted_score >= 80:
        grade = 'B'
        grade_color = '#0dcaf0'
    elif predicted_score >= 70:
        grade = 'C'
        grade_color = '#fd7e14'
    elif predicted_score >= 60:
        grade = 'D'
        grade_color = '#dc3545'
    else:
        grade = 'F'
        grade_color = '#6c757d'

    # Confidence based on data available
    if total_assignments == 0:
        confidence = 'Low'
        confidence_color = '#dc3545'
    elif total_assignments < 3:
        confidence = 'Medium'
        confidence_color = '#fd7e14'
    else:
        confidence = 'High'
        confidence_color = '#198754'

    return {
        'predicted_score': predicted_score,
        'predicted_grade': grade,
        'grade_color': grade_color,
        'confidence': confidence,
        'confidence_color': confidence_color,
        'attendance_pct': round(attendance_pct, 1),
        'assignment_rate': round(assignment_rate, 1),
        'current_score': round(current_score, 1),
    }


# ============================================================================
# 2. DROPOUT RISK DETECTOR
# ============================================================================

def calculate_dropout_risk(attendance_pct, grade_score, assignments_rate,
                            days_since_last_login, courses_failing):
    """
    Calculates dropout risk score (0-100) based on multiple factors.

    Factors:
    - Attendance below 60% → high risk
    - Grade below 50 → high risk
    - Assignment completion below 50% → medium risk
    - Not logged in for 7+ days → medium risk
    - Failing 2+ courses → high risk

    Returns: risk_score, risk_level, risk_color, risk_factors
    """
    risk_score = 0
    risk_factors = []

    # Attendance risk (max 35 points)
    if attendance_pct < 40:
        risk_score += 35
        risk_factors.append('Very low attendance (below 40%)')
    elif attendance_pct < 60:
        risk_score += 25
        risk_factors.append('Low attendance (below 60%)')
    elif attendance_pct < 75:
        risk_score += 10
        risk_factors.append('Attendance below target (below 75%)')

    # Grade risk (max 30 points)
    if grade_score < 40:
        risk_score += 30
        risk_factors.append('Very low grades (below 40%)')
    elif grade_score < 55:
        risk_score += 20
        risk_factors.append('Below average grades (below 55%)')
    elif grade_score < 65:
        risk_score += 10
        risk_factors.append('Grades need improvement (below 65%)')

    # Assignment completion risk (max 20 points)
    if assignments_rate < 30:
        risk_score += 20
        risk_factors.append('Very low assignment submission (below 30%)')
    elif assignments_rate < 60:
        risk_score += 12
        risk_factors.append('Missing many assignments (below 60%)')
    elif assignments_rate < 80:
        risk_score += 5
        risk_factors.append('Some assignments missing (below 80%)')

    # Login activity risk (max 10 points)
    if days_since_last_login > 14:
        risk_score += 10
        risk_factors.append('Inactive for 14+ days')
    elif days_since_last_login > 7:
        risk_score += 5
        risk_factors.append('Inactive for 7+ days')

    # Failing courses risk (max 5 points)
    if courses_failing >= 2:
        risk_score += 5
        risk_factors.append(f'Failing {courses_failing} courses')

    risk_score = min(100, risk_score)

    # Risk level
    if risk_score >= 70:
        risk_level = 'HIGH RISK'
        risk_color = '#dc3545'
        recommendation = 'Immediate intervention needed. Contact student and counselor.'
    elif risk_score >= 40:
        risk_level = 'MEDIUM RISK'
        risk_color = '#fd7e14'
        recommendation = 'Monitor closely. Schedule a check-in with the student.'
    elif risk_score >= 20:
        risk_level = 'LOW RISK'
        risk_color = '#ffc107'
        recommendation = 'Keep an eye on progress. Encourage improvement.'
    else:
        risk_level = 'MINIMAL RISK'
        risk_color = '#198754'
        recommendation = 'Student is doing well. Keep up the good work!'

    return {
        'risk_score': risk_score,
        'risk_level': risk_level,
        'risk_color': risk_color,
        'risk_factors': risk_factors,
        'recommendation': recommendation,
    }


# ============================================================================
# 3. PERSONALIZED STUDY RECOMMENDATIONS
# ============================================================================

from functools import lru_cache

@lru_cache(maxsize=128)
def generate_study_recommendations(attendance_pct, grade_score,
                                    assignment_rate, course_name, api_key=None):
    """
    Generates personalized study recommendations based on student performance.
    Uses Google Gemini if api_key is provided, else falls back to rule-based logic.
    """
    if api_key:
        try:
            import google.generativeai as genai
            import json
            genai.configure(api_key=api_key)
            model = genai.GenerativeModel('gemini-2.5-flash')
            
            prompt = f"""You are an expert academic counselor providing actionable study recommendations for a student taking '{course_name}'.
The student's current performance metrics are:
- Attendance: {attendance_pct}%
- Current Grade/Score: {grade_score}%
- Assignment Completion Rate: {assignment_rate}%

Generate 3 to 4 highly personalized, specific recommendations to help them improve or maintain their performance. 
Return ONLY valid JSON in the exact following structure:
[
  {{
    "priority": "HIGH", 
    "priority_color": "#dc3545", 
    "category": "📅 Attendance", 
    "title": "Short actionable title",
    "description": "Specific reason why this helps based on their metrics.",
    "action": "Concrete next step they should take today."
  }}
]
(Use #dc3545 for HIGH, #fd7e14 for MEDIUM, #0dcaf0 for LOW, #198754 for INFO priority colors).
Do not wrap the response in markdown code blocks, just raw JSON."""
            
            response = model.generate_content(prompt)
            text = response.text.strip()
            
            if text.startswith("```json"): text = text[7:]
            elif text.startswith("```"): text = text[3:]
            if text.endswith("```"): text = text[:-3]
            
            recommendations = json.loads(text.strip())
            
            # Sort by priority
            priority_order = {'HIGH': 0, 'MEDIUM': 1, 'LOW': 2, 'INFO': 3}
            recommendations.sort(key=lambda x: priority_order.get(x.get('priority', 'LOW'), 4))
            
            return recommendations
            
        except Exception as e:
            print(f"[WARN] Gemini AI failed for recommendations, falling back to rule-based: {str(e)}")

    recommendations = []

    # Attendance recommendations
    if attendance_pct < 60:
        recommendations.append({
            'priority': 'HIGH',
            'priority_color': '#dc3545',
            'category': '📅 Attendance',
            'title': 'Improve Attendance Urgently',
            'description': f'Your attendance is {attendance_pct}%. Attend every class for the next 2 weeks to recover.',
            'action': 'Set daily alarms for all classes. Ask a classmate to remind you.',
        })
    elif attendance_pct < 80:
        recommendations.append({
            'priority': 'MEDIUM',
            'priority_color': '#fd7e14',
            'category': '📅 Attendance',
            'title': 'Increase Attendance',
            'description': f'Your attendance is {attendance_pct}%. You need at least 80% to avoid issues.',
            'action': 'Try not to miss more than 1 class per week.',
        })

    # Grade recommendations
    if grade_score < 50:
        recommendations.append({
            'priority': 'HIGH',
            'priority_color': '#dc3545',
            'category': '📝 Grades',
            'title': 'Seek Academic Help Immediately',
            'description': f'Your current score is {grade_score}%. You are at risk of failing.',
            'action': 'Visit your instructor during office hours. Form a study group. Review all past materials.',
        })
    elif grade_score < 65:
        recommendations.append({
            'priority': 'MEDIUM',
            'priority_color': '#fd7e14',
            'category': '📝 Grades',
            'title': 'Improve Your Scores',
            'description': f'Your current score is {grade_score}%. Focus on weak topics.',
            'action': 'Review lecture notes daily. Practice past exam questions.',
        })
    elif grade_score < 80:
        recommendations.append({
            'priority': 'LOW',
            'priority_color': '#0dcaf0',
            'category': '📝 Grades',
            'title': 'Push for Higher Grades',
            'description': f'Your score is {grade_score}%. You are close to achieving a B or above.',
            'action': 'Focus on the hardest topics. Ask for extra practice problems.',
        })

    # Assignment recommendations
    if assignment_rate < 50:
        recommendations.append({
            'priority': 'HIGH',
            'priority_color': '#dc3545',
            'category': '📋 Assignments',
            'title': 'Submit Missing Assignments',
            'description': f'Only {assignment_rate}% of assignments submitted. This heavily impacts your grade.',
            'action': 'List all missing assignments. Submit at least 2 this week. Ask instructor if late submission is allowed.',
        })
    elif assignment_rate < 80:
        recommendations.append({
            'priority': 'MEDIUM',
            'priority_color': '#fd7e14',
            'category': '📋 Assignments',
            'title': 'Complete All Assignments',
            'description': f'You have submitted {assignment_rate}% of assignments.',
            'action': 'Set reminders for assignment deadlines. Complete pending work this weekend.',
        })

    # Study plan based on overall performance
    overall = (attendance_pct * 0.3 + grade_score * 0.5 + assignment_rate * 0.2)
    if overall >= 80:
        recommendations.append({
            'priority': 'INFO',
            'priority_color': '#198754',
            'category': '🎯 Study Plan',
            'title': 'Maintain Excellence',
            'description': 'You are performing well! Keep your momentum going.',
            'action': 'Help classmates, explore advanced topics, prepare early for exams.',
        })
    elif overall >= 60:
        recommendations.append({
            'priority': 'LOW',
            'priority_color': '#0dcaf0',
            'category': '🎯 Study Plan',
            'title': 'Structured Study Schedule',
            'description': 'Create a daily study routine to improve consistently.',
            'action': 'Study 1-2 hours daily. Focus on weak topics first. Review notes before each class.',
        })
    else:
        recommendations.append({
            'priority': 'HIGH',
            'priority_color': '#dc3545',
            'category': '🎯 Study Plan',
            'title': 'Intensive Study Plan Needed',
            'description': 'Your overall performance needs significant improvement.',
            'action': 'Study 3+ hours daily. Use YouTube tutorials. Form a study group. Meet your instructor weekly.',
        })

    # Sort by priority
    priority_order = {'HIGH': 0, 'MEDIUM': 1, 'LOW': 2, 'INFO': 3}
    recommendations.sort(key=lambda x: priority_order.get(x['priority'], 4))

    return recommendations