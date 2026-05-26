import json
import os

import requests

# ============================================
# MISTRAL API CONFIGURATION
# ============================================

API_KEY = "HgQhmaUfQ8Wr6C6gaLA0PFGu3eyPFZqI"

API_URL = "https://api.mistral.ai/v1/chat/completions"

# ============================================
# RESUME ANALYSIS FUNCTION
# ============================================

def local_analysis(resume_text, user_goal, notice):
    resume_lower = resume_text.lower()
    goal_lower = user_goal.lower()
    role_skills = {
        "frontend": ["HTML", "CSS", "JavaScript", "React", "Accessibility"],
        "backend": ["Python", "SQL", "REST API", "Testing", "Docker"],
        "full stack": ["HTML", "CSS", "JavaScript", "Python", "SQL", "REST API"],
        "data": ["Python", "SQL", "Excel", "Power BI", "Statistics"],
        "design": ["Figma", "UX Research", "Prototyping", "Accessibility"],
        "marketing": ["SEO", "Analytics", "Content Strategy", "Campaign Management"],
    }
    general_skills = [
        "Python", "Java", "JavaScript", "TypeScript", "React", "HTML", "CSS",
        "SQL", "Excel", "Power BI", "Figma", "Docker", "AWS", "Git", "SEO"
    ]
    expected = []
    for keyword, skills in role_skills.items():
        if keyword in goal_lower:
            expected = skills
            break
    candidates = expected or general_skills
    skills = [skill for skill in candidates if skill.lower() in resume_lower]
    missing_skills = [skill for skill in expected if skill not in skills]

    return {
        "skills": skills,
        "missing_skills": missing_skills,
        "roadmap": [
            f"Build a practical project demonstrating {skill} for {user_goal} roles."
            for skill in missing_skills[:4]
        ],
        "interview_questions": [
            f"How has your experience prepared you for a {user_goal} role?",
            f"Describe a project where you used {skills[0]}."
            if skills else "Describe a project that demonstrates your strongest skill."
        ],
        "career_options": [],
        "improved_resume": resume_text,
        "analysis_notice": notice
    }


def analyze_resume(resume_text, user_goal):
    if not API_KEY:
        return local_analysis(
            resume_text,
            user_goal,
            "Local analysis shown. Configure MISTRAL_API_KEY to generate an AI-optimized rewrite."
        )

    prompt = f"""
You are an expert resume strategist and hiring manager.

Evaluate and improve the resume based on the user's target role.

USER GOAL:
{user_goal}

STRICT RULES:
- Extract only relevant skills for this goal
- Remove irrelevant technologies
- Identify genuine gaps
- Generate roadmap only for missing skills
- Make output different based on the goal
- Score role alignment from 0 to 100 based on relevance, evidence, clarity, and gaps
- Recommend realistic career options based only on the resume and transferable skills
- Rewrite an improved ATS-friendly resume using only facts present in the source resume
- Never invent employment, education, certifications, projects, metrics, or contact details

Return ONLY valid JSON in this format:

{{
    "skills": [],
    "missing_skills": [],
    "roadmap": [],
    "interview_questions": [],
    "resume_score": 0,
    "score_summary": "",
    "career_options": [
        {{
            "title": "",
            "match_reason": "",
            "skills_to_build": ""
        }}
    ],
    "improved_resume": ""
}}

RESUME:
{resume_text}
"""

    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    data = {
        "model": "mistral-small",
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ]
    }

    try:

        response = requests.post(
            API_URL,
            headers=headers,
            json=data,
            timeout=30
        )
        response.raise_for_status()

        result = response.json()

        content = result["choices"][0]["message"]["content"]

        # Extract JSON safely
        start = content.find("{")
        end = content.rfind("}") + 1
        if start == -1 or end == 0:
            raise ValueError("The AI response did not contain valid JSON.")

        json_text = content[start:end]

        return json.loads(json_text)

    except Exception:
        return local_analysis(
            resume_text,
            user_goal,
            "The AI service is unavailable right now. A local analysis and resume PDF are shown instead."
        )
