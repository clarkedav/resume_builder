from django.shortcuts import render, redirect, get_object_or_404
from django.http import HttpResponse

from .forms import (
    ResumeForm,
    EducationFormSet,
    WorkExperienceFormSet,
    UploadResumeForm,
)
from .models import Resume

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml

import io
import pdfplumber


# ─── ACTION VERBS ──────────────────────────────────────────────────────

ACTION_VERBS = [
    # Leadership & Management
    'administered',
    'appointed',
    'coached',
    'coordinated',
    'delegated',
    'directed',
    'facilitated',
    'governed',
    'guided',
    'managed',
    'mentored',
    'mobilized',
    'organized',
    'oversaw',
    'prioritized',
    'regulated',
    'resolved',
    'steered',
    'supervised',

    # Achievement & Results
    'accomplished',
    'achieved',
    'advanced',
    'attained',
    'boosted',
    'completed',
    'exceeded',
    'expanded',
    'generated',
    'grew',
    'maximized',
    'outperformed',
    'reached',
    'recovered',
    'strengthened',
    'surpassed',
    'won',

    # Technology & Engineering
    'architected',
    'automated',
    'built',
    'coded',
    'configured',
    'debugged',
    'deployed',
    'designed',
    'developed',
    'engineered',
    'implemented',
    'integrated',
    'migrated',
    'programmed',
    'refactored',
    'tested',
    'troubleshot',
    'upgraded',
    'validated',

    # Data & Analytics
    'analyzed',
    'assessed',
    'audited',
    'calculated',
    'compiled',
    'correlated',
    'evaluated',
    'forecasted',
    'identified',
    'interpreted',
    'measured',
    'modeled',
    'monitored',
    'quantified',
    'researched',
    'segmented',
    'tracked',
    'visualized',

    # Project & Operations
    'allocated',
    'assembled',
    'centralized',
    'consolidated',
    'executed',
    'operationalized',
    'planned',
    'processed',
    'scheduled',
    'standardized',
    'streamlined',
    'structured',

    # Communication & Collaboration
    'advocated',
    'authored',
    'briefed',
    'collaborated',
    'communicated',
    'consulted',
    'convinced',
    'corresponded',
    'liaised',
    'negotiated',
    'partnered',
    'presented',
    'proposed',
    'promoted',
    'published',
    'reported',
    'translated',

    # Sales & Business Development
    'acquired',
    'captured',
    'closed',
    'converted',
    'cultivated',
    'prospected',
    'retained',
    'renewed',
    'secured',
    'sold',
    'upsold',

    # Marketing
    'branded',
    'campaigned',
    'conceptualized',
    'curated',
    'launched',
    'marketed',
    'positioned',
    'targeted',

    # Finance & Administration
    'accounted',
    'budgeted',
    'controlled',
    'reconciled',
    'verified',

    # Education & Training
    'assessed',
    'coached',
    'developed',
    'educated',
    'evaluated',
    'facilitated',
    'guided',
    'instructed',
    'mentored',
    'presented',
    'trained',

    # Innovation & Problem Solving
    'accelerated',
    'customized',
    'diagnosed',
    'discovered',
    'devised',
    'formulated',
    'initiated',
    'innovated',
    'modernized',
    'originated',
    'personalized',
    'pioneered',
    'refined',
    'remediated',
    'revitalized',
    'spearheaded',
    'transformed',
    'uncovered',
]


# ─── HOME PAGE ─────────────────────────────────────────────────────────

def home(request):
    """
    Display the landing page for the resume builder application.
    
    This view renders the home page which introduces users to the 
    resume builder and upload features with call-to-action buttons.
    
    Args:
        request (HttpRequest): The HTTP request object
    
    Returns:
        HttpResponse: Rendered home page template (resume/index.html)
    """
    return render(request, 'resume/index.html')


# ─── STATUS HELPER ────────────────────────────────────────────────────

def get_status(score, max_score):
    """
    Determine checkpoint status based on score percentage.
    
    Converts a numeric score to a status label:
    - 100% (full score) = 'good' (green)
    - 50%-99% (partial) = 'in-progress' (yellow)
    - 0%-49% (failing) = 'needs-work' (red)
    
    Args:
        score (int): The numeric score achieved
        max_score (int): The maximum possible score for the checkpoint
    
    Returns:
        str: One of 'good', 'in-progress', or 'needs-work'
    """
    if score == max_score:
        return 'good'
    elif score >= (max_score * 0.5):
        return 'in-progress'
    else:
        return 'needs-work'


# ─── SCORING HELPER ───────────────────────────────────────────────────

def score_resume_data(
    summary,
    skills_list,
    experiences,
    educations,
    linkedin,
    target_position
):
    """
    Calculate a comprehensive resume score using 8 checkpoints.
    
    Evaluates a resume across 8 distinct criteria:
    1. Action Verbs - use of strong power verbs (10 pts)
    2. Quantifiable Results - specific numbers and metrics (10 pts)
    3. Experience Detail - comprehensive job descriptions (10 pts)
    4. Role-Specific Keywords - targeted skills for position (10 pts)
    5. Professional Summary - compelling career overview (10 pts)
    6. LinkedIn & Contact - complete contact information (10 pts)
    7. Skills List - sufficient relevant skills (10 pts)
    8. Education & Accomplishments - degrees and achievements (10 pts)
    
    Maximum total score: 80 points (100%)
    
    Args:
        summary (str): Professional summary text from resume
        skills_list (list): List of skills entered by user
        experiences (QuerySet): Work experience entries from database
        educations (QuerySet): Education entries from database
        linkedin (str): LinkedIn profile URL
        target_position (str): Target job position for keyword matching
    
    Returns:
        dict: Dictionary containing:
            - total_score (int): Sum of all checkpoint scores (0-80)
            - max_score (int): Maximum possible score (80)
            - percentage (int): Percentage score (0-100)
            - checkpoints (list): List of dicts with individual checkpoint results
    """
    checkpoints = []
    total_score = 0

    # Handle None values to prevent errors
    summary = summary or ''
    skills_list = skills_list or []
    experiences = experiences or []
    educations = educations or []
    target_position = target_position or ''
    linkedin = linkedin or ''

    # ─── CHECKPOINT 1: Action Verbs ────────────────────────────────────

    action_verbs_count = 0

    if experiences:
        all_desc = ' '.join(
            (experience.description or '').lower()
            for experience in experiences
        )

        action_verbs_count = sum(
            1
            for verb in ACTION_VERBS
            if verb in all_desc
        )

    if action_verbs_count >= 5:
        action_verbs_score = 10
        action_verbs_feedback = (
            'You used strong action verbs effectively '
            'throughout your descriptions.'
        )
    elif action_verbs_count >= 3:
        action_verbs_score = 6
        action_verbs_feedback = (
            'You used some action verbs. Aim for at least '
            '5 different action verbs such as Developed, Led, '
            'Managed, and Implemented.'
        )
    else:
        action_verbs_score = 0
        action_verbs_feedback = (
            'Start each experience bullet with a strong action verb. '
            'Avoid weak verbs such as Worked, Helped, or Did.'
        )

    action_verbs_tip = (
        'Examples: Developed, Led, Managed, Implemented, '
        'Increased, Designed, Delivered, Coordinated'
    )

    checkpoints.append({
        'name': 'Action Verbs',
        'icon': '⚡',
        'score': action_verbs_score,
        'max': 10,
        'status': get_status(action_verbs_score, 10),
        'feedback': action_verbs_feedback,
        'tip': action_verbs_tip,
    })

    total_score += action_verbs_score

    # ─── CHECKPOINT 2: Quantifiable Results ────────────────────────────

    has_numbers = False

    if experiences:
        all_desc = ' '.join(
            (experience.description or '').lower()
            for experience in experiences
        )

        has_numbers = any(
            character.isdigit()
            for character in all_desc
        )

    if has_numbers:
        quantifiable_score = 10
        quantifiable_feedback = (
            'Your descriptions include specific numbers '
            'and measurable results.'
        )
    else:
        quantifiable_score = 0
        quantifiable_feedback = (
            'Add numbers to show impact. Include percentages '
            '(40%), quantities (500+ users), revenue, savings, '
            'or team sizes.'
        )

    quantifiable_tip = (
        'Examples: Increased sales by 35%, '
        'Managed team of 8, Served 500+ clients'
    )

    checkpoints.append({
        'name': 'Quantifiable Results',
        'icon': '📊',
        'score': quantifiable_score,
        'max': 10,
        'status': get_status(quantifiable_score, 10),
        'feedback': quantifiable_feedback,
        'tip': quantifiable_tip,
    })

    total_score += quantifiable_score

    # ─── CHECKPOINT 3: Experience Detail ──────────────────────────────

    if experiences:
        detailed = all(
            len(experience.description or '') >= 120
            for experience in experiences
        )
    else:
        detailed = False

    if detailed and len(experiences) >= 2:
        experience_score = 10
        experience_feedback = (
            'Your experience descriptions are detailed '
            'and well-articulated.'
        )
    elif detailed or len(experiences) >= 2:
        experience_score = 6
        experience_feedback = (
            'Expand your descriptions to 2-3 sentences per role. '
            'Include what you did, how you did it, and the impact.'
        )
    else:
        experience_score = 0
        experience_feedback = (
            'Add at least 2 work experiences with detailed '
            'descriptions of your responsibilities and achievements.'
        )

    experience_tip = (
        'Each role should explain what you did, how you did it, '
        'and the measurable outcome.'
    )

    checkpoints.append({
        'name': 'Experience Detail',
        'icon': '💼',
        'score': experience_score,
        'max': 10,
        'status': get_status(experience_score, 10),
        'feedback': experience_feedback,
        'tip': experience_tip,
    })

    total_score += experience_score

    # ─── CHECKPOINT 4: Job-Specific Keywords ──────────────────────────

    job_keywords_dict = {
        'software': [
            'python',
            'java',
            'javascript',
            'sql',
            'database',
            'api',
            'rest',
            'git',
            'django',
            'react',
        ],
        'data': [
            'python',
            'sql',
            'analytics',
            'tableau',
            'excel',
            'analysis',
            'statistics',
        ],
        'marketing': [
            'campaign',
            'seo',
            'social',
            'content',
            'analytics',
            'engagement',
            'branding',
        ],
        'sales': [
            'closed',
            'revenue',
            'pipeline',
            'client',
            'negotiation',
            'quota',
        ],
        'project': [
            'managed',
            'coordinated',
            'timeline',
            'delivered',
            'stakeholders',
            'agile',
        ],
    }

    relevant_keywords = []
    target_lower = target_position.lower()

    for category, keywords in job_keywords_dict.items():
        if category in target_lower:
            relevant_keywords = keywords
            break

    if relevant_keywords and experiences:
        all_desc = ' '.join(
            (experience.description or '').lower()
            for experience in experiences
        )

        found_keywords = [
            keyword
            for keyword in relevant_keywords
            if keyword in all_desc
        ]

        if len(found_keywords) >= 3:
            keyword_score = 10
            keyword_feedback = (
                f'Your resume emphasizes key skills '
                f'for a {target_position} role.'
            )
        elif len(found_keywords) >= 1:
            keyword_score = 6
            missing = [
                keyword
                for keyword in relevant_keywords
                if keyword not in all_desc
            ][:2]

            if missing:
                keyword_feedback = (
                    f'Emphasize more {target_position}-specific skills: '
                    f'{", ".join(missing)}'
                )
            else:
                keyword_feedback = (
                    f'Add more {target_position}-specific skills '
                    'to strengthen your resume.'
                )
        else:
            keyword_score = 0
            keyword_feedback = (
                f'Add {target_position}-specific keywords. '
                f'Missing: {", ".join(relevant_keywords[:3])}'
            )
    elif target_position:
        keyword_score = 5
        keyword_feedback = (
            f'Add experience entries that highlight '
            f'{target_position}-specific skills.'
        )
    else:
        keyword_score = 10
        keyword_feedback = (
            'No specific target position selected, '
            'so this checkpoint is not applicable.'
        )

    keyword_tip = (
        f'For {target_position or "your target role"}: '
        'highlight the most relevant technical and professional skills.'
    )

    checkpoints.append({
        'name': 'Role-Specific Skills',
        'icon': '🎯',
        'score': keyword_score,
        'max': 10,
        'status': get_status(keyword_score, 10),
        'feedback': keyword_feedback,
        'tip': keyword_tip,
    })

    total_score += keyword_score

    # ─── CHECKPOINT 5: Professional Summary ───────────────────────────

    summary_length = len(summary.strip())

    if summary_length >= 150:
        summary_score = 10
        summary_feedback = (
            'Your summary is detailed and clearly '
            'communicates your professional value.'
        )
    elif summary_length >= 80:
        summary_score = 6
        summary_feedback = (
            'Strengthen your summary. Expand it to 3-4 sentences '
            'and include a specific achievement.'
        )
    else:
        summary_score = 0
        summary_feedback = (
            'Write a professional summary of 3-4 sentences. '
            'Highlight who you are, your key skills, and achievements.'
        )

    summary_tip = (
        'Format: Who you are + Key skill + '
        'Concrete achievement + Career goal'
    )

    checkpoints.append({
        'name': 'Professional Summary',
        'icon': '📝',
        'score': summary_score,
        'max': 10,
        'status': get_status(summary_score, 10),
        'feedback': summary_feedback,
        'tip': summary_tip,
    })

    total_score += summary_score

    # ─── CHECKPOINT 6: LinkedIn & Contact Info ─────────────────────────

    if linkedin and summary.strip():
        linkedin_score = 10
        linkedin_feedback = (
            'Your LinkedIn URL and professional summary '
            'make your profile more complete.'
        )
    elif linkedin or summary.strip():
        linkedin_score = 5
        linkedin_feedback = (
            'Add both a professional summary and LinkedIn URL '
            'to strengthen your professional presence.'
        )
    else:
        linkedin_score = 0
        linkedin_feedback = (
            'Add your LinkedIn profile URL and professional '
            'summary for a more complete resume.'
        )

    linkedin_tip = (
        'Include Email, Phone, Location, LinkedIn URL, '
        'and an optional personal website.'
    )

    checkpoints.append({
        'name': 'LinkedIn & Contact',
        'icon': '🔗',
        'score': linkedin_score,
        'max': 10,
        'status': get_status(linkedin_score, 10),
        'feedback': linkedin_feedback,
        'tip': linkedin_tip,
    })

    total_score += linkedin_score

    # ─── CHECKPOINT 7: Skills List ────────────────────────────────────

    skills_count = len(skills_list)

    if skills_count >= 10:
        skills_score = 10
        skills_feedback = (
            'Your skills list is comprehensive and well-rounded.'
        )
    elif skills_count >= 7:
        skills_score = 7
        skills_feedback = (
            f'You have {skills_count} skills. '
            'Add 3-4 more for a comprehensive list.'
        )
    elif skills_count >= 5:
        skills_score = 4
        skills_feedback = (
            f'You have {skills_count} skills. '
            'Aim for at least 10 relevant skills.'
        )
    else:
        skills_score = 0
        skills_feedback = (
            'Add at least 8-10 relevant skills. Include technical '
            'skills and important professional skills.'
        )

    skills_tip = (
        'Mix technical skills with relevant professional skills. '
        'Focus on skills related to your target role.'
    )

    checkpoints.append({
        'name': 'Skills List',
        'icon': '⚙️',
        'score': skills_score,
        'max': 10,
        'status': get_status(skills_score, 10),
        'feedback': skills_feedback,
        'tip': skills_tip,
    })

    total_score += skills_score

    # ─── CHECKPOINT 8: Education & Credentials + Accomplishments ──────

    education_score = 0
    education_feedback = ''

    if educations:
        # Count complete education records
        complete_educations = sum(
            1
            for education in educations
            if (
                education.degree
                and education.institution
                and education.grad_year
            )
        )

        # Check for GPA in any education entry
        has_gpa = any(
            education.gpa
            for education in educations
        )

        # Count accomplishments with action verbs
        accomplishments_quality = 0
        total_accomplishments = 0

        for education in educations:
            accomplishments_text = (
                education.accomplishments or ''
            )

            accomplishments = [
                item.strip()
                for item in accomplishments_text.split(',')
                if item.strip()
            ]

            total_accomplishments += len(accomplishments)

            for accomplishment in accomplishments:
                words = accomplishment.split()

                if not words:
                    continue

                first_word = (
                    words[0]
                    .lower()
                    .strip('.,:;!?()[]{}')
                )

                if first_word in ACTION_VERBS:
                    accomplishments_quality += 1

        has_quality_accomplishments = (
            accomplishments_quality >= 2
        )

        # Score based on completeness
        if complete_educations == len(educations):
            if has_gpa and has_quality_accomplishments:
                education_score = 10
                education_feedback = (
                    'Your education entries are complete and include '
                    'GPA plus accomplishments that begin with strong '
                    'action verbs.'
                )
            elif has_gpa or has_quality_accomplishments:
                education_score = 8

                if has_gpa:
                    education_feedback = (
                        'Your education details are complete and include '
                        'a GPA. Add 2+ accomplishments beginning with '
                        'strong action verbs for full credit.'
                    )
                else:
                    education_feedback = (
                        'Your education details are complete and include '
                        'quality accomplishments. Add a GPA if it is '
                        'strong and relevant to your target role.'
                    )
            elif total_accomplishments > 0:
                education_score = 7
                education_feedback = (
                    'Your education details are complete, but your '
                    'accomplishments should begin with strong action '
                    'verbs such as Achieved, Led, Developed, or Organized.'
                )
            else:
                education_score = 6
                education_feedback = (
                    'Your education details are complete. Add 2+ '
                    'accomplishments using strong action verbs to '
                    'show leadership, achievement, or project work.'
                )
        else:
            education_score = 4
            education_feedback = (
                'Some education details are missing. Make sure every '
                'entry includes degree, institution, and graduation year.'
            )
    else:
        education_score = 0
        education_feedback = (
            'No education information provided. Add your degree, '
            'institution, graduation year, and relevant accomplishments.'
        )

    education_tip = (
        'For each education entry, include Degree, Institution, '
        'Graduation Year, GPA when appropriate, and 2+ '
        'accomplishments beginning with action verbs such as '
        'Achieved, Led, Developed, Organized, or Coordinated.'
    )

    checkpoints.append({
        'name': 'Education & Credentials',
        'icon': '🎓',
        'score': education_score,
        'max': 10,
        'status': get_status(education_score, 10),
        'feedback': education_feedback,
        'tip': education_tip,
    })

    total_score += education_score

    # ─── RETURN SCORE ─────────────────────────────────────────────────

    return {
        'total_score': total_score,
        'max_score': 80,
        'percentage': int((total_score / 80) * 100),
        'checkpoints': checkpoints,
    }


# ─── FORM & DISPLAY VIEWS ──────────────────────────────────────────────

def resume_form(request):
    """
    Handle resume creation with multi-part form including formsets.
    
    Manages a complex form with:
    - Personal resume information (ResumeForm)
    - Multiple education entries (EducationFormSet)
    - Multiple work experience entries (WorkExperienceFormSet)
    
    On GET: Display blank form with formsets ready for input
    On POST: Validate and save all form data to database, redirect to display
    
    Args:
        request (HttpRequest): The HTTP request object
    
    Returns:
        HttpResponse: Rendered form template on GET or after validation error
                     Redirect to resume_display on successful POST
    """
    if request.method == 'POST':
        form = ResumeForm(request.POST)
        edu_formset = EducationFormSet(request.POST, prefix='edu')
        exp_formset = WorkExperienceFormSet(request.POST, prefix='exp')

        if (
            form.is_valid()
            and edu_formset.is_valid()
            and exp_formset.is_valid()
        ):
            resume = form.save()
            edu_formset.instance = resume
            edu_formset.save()
            exp_formset.instance = resume
            exp_formset.save()

            return redirect('resume_display', pk=resume.pk)

    else:
        form = ResumeForm()
        edu_formset = EducationFormSet(prefix='edu')
        exp_formset = WorkExperienceFormSet(prefix='exp')

    return render(
        request,
        'resume/form.html',
        {
            'form': form,
            'edu_formset': edu_formset,
            'exp_formset': exp_formset,
        }
    )


# ─── RESUME DISPLAY ────────────────────────────────────────────────────

def resume_display(request, pk):
    """
    Display a formatted resume preview in professional layout.
    
    Retrieves a resume and all related data (education, experiences, skills)
    and renders them in a polished, printable format. Accomplishments are
    parsed from comma-separated text into lists for display.
    
    Args:
        request (HttpRequest): The HTTP request object
        pk (int): Primary key of the Resume object to display
    
    Returns:
        HttpResponse: Rendered resume display page
        Http404: If resume with given pk does not exist
    """
    resume = get_object_or_404(Resume, pk=pk)

    educations = resume.educations.all()
    experiences = resume.experiences.all()

    skills_list = [
        skill.strip()
        for skill in (resume.skills or '').split(',')
        if skill.strip()
    ]

    for education in educations:
        education.accomplishments_list = [
            accomplishment.strip()
            for accomplishment in (
                education.accomplishments or ''
            ).split(',')
            if accomplishment.strip()
        ]

    return render(
        request,
        'resume/display.html',
        {
            'resume': resume,
            'educations': educations,
            'experiences': experiences,
            'skills_list': skills_list,
        }
    )


# ─── RESUME FEEDBACK ──────────────────────────────────────────────────

def resume_feedback(request, pk):
    """
    Display the comprehensive 8-point resume score report.
    
    Calculates resume score using score_resume_data(), generates detailed
    feedback for all 8 checkpoints, and displays results with visual
    indicators (green/yellow/red status badges).
    
    Args:
        request (HttpRequest): The HTTP request object
        pk (int): Primary key of the Resume object to score
    
    Returns:
        HttpResponse: Rendered feedback/score report page
        Http404: If resume with given pk does not exist
    """
    resume = get_object_or_404(Resume, pk=pk)

    educations = resume.educations.all()
    experiences = resume.experiences.all()

    skills_list = [
        skill.strip()
        for skill in (resume.skills or '').split(',')
        if skill.strip()
    ]

    scores = score_resume_data(
        summary=resume.summary or '',
        skills_list=skills_list,
        experiences=experiences,
        educations=educations,
        linkedin=resume.linkedin or '',
        target_position=resume.target_position or '',
    )

    return render(
        request,
        'resume/feedback.html',
        {
            'resume': resume,
            **scores,
        }
    )


# ─── DOWNLOAD DOCX ─────────────────────────────────────────────────────

def download_docx(request, pk):
    """
    Generate and download resume as a formatted DOCX file.
    
    Creates a professional Word document with:
    - Letter size (8.5" x 11") with standard margins
    - Colored section headers (purple gradient backgrounds)
    - Skills in 4-column grid layout
    - Proper spacing and typography for one-page format
    - Support for accomplishments with bullet points
    
    Args:
        request (HttpRequest): The HTTP request object
        pk (int): Primary key of the Resume object to download
    
    Returns:
        HttpResponse: DOCX file as attachment download
        Http404: If resume with given pk does not exist
    """
    resume = get_object_or_404(Resume, pk=pk)

    educations = resume.educations.all()
    experiences = resume.experiences.all()

    skills_list = [
        skill.strip()
        for skill in (resume.skills or '').split(',')
        if skill.strip()
    ]

    doc = Document()

    # ─── Page setup ────────────────────────────────────────────────────

    for section in doc.sections:
        section.page_height = Inches(11)
        section.page_width = Inches(8.5)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    # ─── Name ──────────────────────────────────────────────────────────

    name_para = doc.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    name_para.paragraph_format.space_before = Pt(0)
    name_para.paragraph_format.space_after = Pt(2)
    name_para.paragraph_format.line_spacing = 1.0

    name_run = name_para.add_run(resume.full_name or '')
    name_run.bold = True
    name_run.font.size = Pt(14)
    name_run.font.color.rgb = RGBColor(102, 126, 234)

    # ─── Target position ───────────────────────────────────────────────

    if resume.target_position:
        position_para = doc.add_paragraph()
        position_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        position_para.paragraph_format.space_before = Pt(0)
        position_para.paragraph_format.space_after = Pt(1)
        position_para.paragraph_format.line_spacing = 1.0

        position_run = position_para.add_run(resume.target_position)
        position_run.bold = True
        position_run.font.size = Pt(10)
        position_run.font.color.rgb = RGBColor(118, 75, 162)

    # ─── Contact information ───────────────────────────────────────────

    contact_parts = []

    if resume.email:
        contact_parts.append(resume.email)
    if resume.phone:
        contact_parts.append(resume.phone)
    if resume.location:
        contact_parts.append(resume.location)
    if resume.linkedin:
        contact_parts.append(resume.linkedin)

    if contact_parts:
        contact_para = doc.add_paragraph(' | '.join(contact_parts))
        contact_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        contact_para.paragraph_format.space_before = Pt(0)
        contact_para.paragraph_format.space_after = Pt(3)
        contact_para.paragraph_format.line_spacing = 1.0

        for run in contact_para.runs:
            run.font.size = Pt(8)

    # ─── Section title helper ──────────────────────────────────────────

    def add_section_title(title):
        """
        Add a formatted section header with purple background and border.
        
        Creates a styled header paragraph with:
        - Purple background color (#6B7EEA)
        - White text
        - Bottom border in matching color
        
        Args:
            title (str): The section header text (e.g., 'SKILLS', 'EXPERIENCE')
        
        Returns:
            None (modifies document in place)
        """
        header = doc.add_paragraph()
        header.paragraph_format.space_before = Pt(0)
        header.paragraph_format.space_after = Pt(2)
        header.paragraph_format.line_spacing = 1.0

        header_run = header.add_run(title)
        header_run.bold = True
        header_run.font.size = Pt(9)
        header_run.font.color.rgb = RGBColor(255, 255, 255)

        shading = parse_xml(
            '<w:shd '
            'xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main" '
            'w:fill="6B7EEA"/>'
        )

        header._element.get_or_add_pPr().append(shading)

        paragraph_properties = (
            header._element.get_or_add_pPr()
        )

        border = parse_xml(
            '<w:pBdr '
            'xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main">'
            '<w:bottom '
            'w:val="single" '
            'w:sz="12" '
            'w:space="1" '
            'w:color="6B7EEA"/>'
            '</w:pBdr>'
        )

        paragraph_properties.append(border)

    # ─── Summary ───────────────────────────────────────────────────────

    if resume.summary:
        add_section_title('PROFESSIONAL SUMMARY')

        summary_para = doc.add_paragraph(resume.summary)
        summary_para.paragraph_format.space_before = Pt(0)
        summary_para.paragraph_format.space_after = Pt(2)
        summary_para.paragraph_format.line_spacing = 1.15

        for run in summary_para.runs:
            run.font.size = Pt(9)

    # ─── Skills ────────────────────────────────────────────────────────

    if skills_list:
        add_section_title('SKILLS')

        columns = 4
        rows = (len(skills_list) + columns - 1) // columns

        table = doc.add_table(rows=rows, cols=columns)

        table_borders = parse_xml(
            '<w:tblBorders '
            'xmlns:w="http://schemas.openxmlformats.org/'
            'wordprocessingml/2006/main">'
            '<w:top w:val="none"/>'
            '<w:left w:val="none"/>'
            '<w:bottom w:val="none"/>'
            '<w:right w:val="none"/>'
            '<w:insideH w:val="none"/>'
            '<w:insideV w:val="none"/>'
            '</w:tblBorders>'
        )

        table._tbl.tblPr.append(table_borders)

        skill_index = 0

        for row_index in range(rows):
            for column_index in range(columns):
                if skill_index < len(skills_list):
                    cell = table.rows[row_index].cells[column_index]
                    cell.text = ''

                    paragraph = cell.paragraphs[0]
                    paragraph.style = 'List Bullet'
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)

                    run = paragraph.add_run(skills_list[skill_index])
                    run.font.size = Pt(9)

                    skill_index += 1

    # ─── Work Experience ───────────────────────────────────────────────

    if experiences:
        add_section_title('WORK EXPERIENCE')

        for experience in experiences:
            job_para = doc.add_paragraph()
            job_para.paragraph_format.space_before = Pt(0)
            job_para.paragraph_format.space_after = Pt(1)

            job_run = job_para.add_run(experience.job_title or '')
            job_run.bold = True
            job_run.font.size = Pt(10)
            job_run.font.color.rgb = RGBColor(102, 126, 234)

            if experience.company:
                company_para = doc.add_paragraph(experience.company)
                company_para.paragraph_format.space_before = Pt(0)
                company_para.paragraph_format.space_after = Pt(1)

                for run in company_para.runs:
                    run.font.size = Pt(9)

            details = []

            if experience.duration:
                details.append(experience.duration)
            if experience.work_location:
                details.append(experience.work_location)

            if details:
                details_para = doc.add_paragraph(' | '.join(details))
                details_para.paragraph_format.space_before = Pt(0)
                details_para.paragraph_format.space_after = Pt(1)

                for run in details_para.runs:
                    run.font.size = Pt(8)
                    run.font.italic = True

            if experience.description:
                description_para = doc.add_paragraph(
                    experience.description
                )

                description_para.paragraph_format.space_before = Pt(0)
                description_para.paragraph_format.space_after = Pt(2)

                for run in description_para.runs:
                    run.font.size = Pt(9)

    # ─── Education ─────────────────────────────────────────────────────

    if educations:
        add_section_title('EDUCATION')

        for education in educations:
            degree_para = doc.add_paragraph()
            degree_para.paragraph_format.space_before = Pt(0)
            degree_para.paragraph_format.space_after = Pt(1)

            degree_run = degree_para.add_run(education.degree or '')
            degree_run.bold = True
            degree_run.font.size = Pt(10)
            degree_run.font.color.rgb = RGBColor(102, 126, 234)

            if education.institution:
                institution_para = doc.add_paragraph(
                    education.institution
                )

                institution_para.paragraph_format.space_before = Pt(0)
                institution_para.paragraph_format.space_after = Pt(1)

                for run in institution_para.runs:
                    run.font.size = Pt(9)

            details = []

            if education.grad_year:
                details.append(str(education.grad_year))
            if education.gpa:
                details.append(f'GPA: {education.gpa}')

            if details:
                details_para = doc.add_paragraph(' | '.join(details))
                details_para.paragraph_format.space_before = Pt(0)
                details_para.paragraph_format.space_after = Pt(1)

                for run in details_para.runs:
                    run.font.size = Pt(8)
                    run.font.italic = True

            # ─── Education accomplishments ────────────────────────────

            if education.accomplishments:
                accomplishments = [
                    item.strip()
                    for item in (
                        education.accomplishments or ''
                    ).split(',')
                    if item.strip()
                ]

                for accomplishment in accomplishments:
                    accomplishment_para = doc.add_paragraph(
                        accomplishment,
                        style='List Bullet'
                    )

                    accomplishment_para.paragraph_format.space_before = Pt(0)
                    accomplishment_para.paragraph_format.space_after = Pt(0)

                    for run in accomplishment_para.runs:
                        run.font.size = Pt(9)

    # ─── Create DOCX response ──────────────────────────────────────────

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    response = HttpResponse(
        buffer.getvalue(),
        content_type=(
            'application/vnd.openxmlformats-officedocument.'
            'wordprocessingml.document'
        )
    )

    filename = (
        f'{resume.full_name or "Resume"}_Resume.docx'
    )

    response[
        'Content-Disposition'
    ] = f'attachment; filename="{filename}"'

    return response


# ─── UPLOAD & SCAN VIEW ────────────────────────────────────────────────

def upload_scan(request):
    """
    Handle resume file upload and analyze with scoring system.
    
    Accepts DOCX or PDF files, extracts text content, parses basic
    resume information (skills, LinkedIn), and calculates a score
    using the same 8-point evaluation system as the builder.
    
    Supported formats:
    - DOCX: Uses python-docx library
    - PDF: Uses pdfplumber library
    
    Args:
        request (HttpRequest): The HTTP request object
    
    Returns:
        HttpResponse: 
            - On GET: Rendered upload form
            - On POST (success): Rendered scan results page with score
            - On POST (invalid): Upload form with error message
        Http404: If resume with given pk does not exist
    """
    if request.method == 'POST':
        form = UploadResumeForm(request.POST, request.FILES)

        if form.is_valid():
            uploaded_file = form.cleaned_data['resume_file']

            target_position = (
                form.cleaned_data.get('target_position', '')
                or ''
            )

            text = ''
            filename_lower = (
                uploaded_file.name.lower()
            )

            # ─── DOCX extraction ──────────────────────────────────────

            if filename_lower.endswith('.docx'):
                document = Document(uploaded_file)

                paragraphs = [
                    paragraph.text
                    for paragraph in document.paragraphs
                    if paragraph.text.strip()
                ]

                text = '\n'.join(paragraphs)

            # ─── PDF extraction ───────────────────────────────────────

            elif filename_lower.endswith('.pdf'):
                with pdfplumber.open(uploaded_file) as pdf:
                    extracted_pages = []

                    for page in pdf.pages:
                        extracted = (
                            page.extract_text()
                        )

                        if extracted:
                            extracted_pages.append(
                                extracted
                            )

                    text = '\n'.join(
                        extracted_pages
                    )

            # ─── Unsupported file ─────────────────────────────────────

            else:
                return render(
                    request,
                    'resume/upload.html',
                    {
                        'form': form,
                        'error': (
                            'Please upload a DOCX or PDF file.'
                        ),
                    }
                )

            # ─── Parse extracted text ──────────────────────────────────

            text_lower = text.lower()

            # Find a basic skills section
            skills_text = ''

            if 'skill' in text_lower:
                skills_index = text_lower.find(
                    'skill'
                )

                skills_text = text[
                    skills_index:skills_index + 500
                ]

            skills_list = [
                skill.strip()
                for skill in skills_text.split(',')
                if skill.strip()
            ][:15]

            # Detect LinkedIn
            has_linkedin = (
                'linkedin.com' in text_lower
                or 'linkedin' in text_lower
            )

            # ─── Score uploaded resume ─────────────────────────────────

            scores = score_resume_data(
                summary=text[:300],
                skills_list=skills_list,
                experiences=[],
                educations=[],
                linkedin=has_linkedin,
                target_position=target_position,
            )

            resume_name = (
                uploaded_file.name
                .rsplit('.', 1)[0]
            )

            return render(
                request,
                'resume/scan_result.html',
                {
                    'resume_name': resume_name,
                    'target_position': target_position,
                    'text_preview': (
                        text[:500] + '...'
                        if len(text) > 500
                        else text
                    ),
                    **scores,
                }
            )

    else:
        form = UploadResumeForm()

    return render(
        request,
        'resume/upload.html',
        {
            'form': form
        }
    )