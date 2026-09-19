MyResumeBuilder - Resume Creation & Scoring Platform
Overview
MyResumeBuilder is a Django web application designed to help job seekers create professional, ATS-optimized resumes and receive intelligent feedback on their resume quality. As a software engineer, I built this project to deepen my full-stack development skills, particularly in backend data processing, database design, and creating intuitive user interfaces that translate complex scoring logic into actionable insights.

To start the application:

Clone the repository and navigate to the project directory
Install dependencies: pip install -r requirements.txt
Run migrations: python manage.py makemigrations && python manage.py migrate
Start the development server: python manage.py runserver
Open your browser to http://127.0.0.1:8000/ to see the home page
The application guides users through an interactive resume builder, evaluates their resume against 8 key professional standards, and generates a detailed score report with actionable recommendations. I created this to solve a real problem: most job seekers lack structured feedback on resume quality, and this tool provides that guidance while also demonstrating how to build a data-driven web application with meaningful user feedback loops.

[Software Demo Video](https://youtu.be/tmajSMEcq9E)

Web Pages
Home Page — The landing page introduces users to the builder and upload features. It displays call-to-action buttons for either building a new resume from scratch or uploading an existing resume for analysis. This page is dynamically generated with links to create new resume sessions.

Resume Builder Form — A multi-section form where users input resume information across five main areas: professional summary, skills, work experience, education, and contact details. The page uses formsets to allow users to add/remove multiple work experiences and education entries dynamically. Accomplishments can be added to education entries. Form validation ensures data quality before submission.

Resume Display — After submission, users see a beautifully formatted preview of their resume matching professional standards. The page displays the resume in a colorful, polished layout with purple gradient section headers and pink gradient skill tags. It includes buttons to print to PDF, download as DOCX, or view the score report. This page dynamically renders all entered data into a professional document format.

Score Report Page — The core feedback feature showing an 8-point evaluation of the resume. It displays a progress score banner (name/target position), a grid of color-coded checkpoint cards (Green = Excellent/10, Yellow = Needs Work/5-9, Red = Needs Help/0-4), detailed feedback for each checkpoint, and actionable tips. The page dynamically calculates scores based on action verbs, quantifiable results, keywords, summary quality, LinkedIn presence, skills count, experience detail, and education completeness.

Resume Upload & Scan — Users can upload existing resumes in DOCX or PDF format. The application extracts text, analyzes it using the same 8-point scoring system, and displays results on an identical results page. This allows users to get feedback on existing resumes before rebuilding them.

The application follows a clear flow: Home → Build/Upload → View Resume → View Score → Edit & Improve.

# Development Environment
Tools & Framework:

Django 6.1.1 (Python web framework)
SQLite3 (database)
Python 3.14.6
VS Code (code editor)
Git (version control)
Libraries & Dependencies:

django-crispy-forms (form rendering)
python-docx (DOCX file generation with colored headers and formatting)
pdfplumber (PDF text extraction for resume scanning)
Pillow (image handling)
Frontend:

HTML5, CSS3 (responsive design)
JavaScript (dynamic form fields, interactivity)
CSS Gradients & animations (UI visual appeal)

# Useful Websites

* [Django Official Documentation](https://docs.djangoproject.com/)
* [Django Forms & Formsets](https://docs.djangoproject.com/en/stable/topics/forms/formsets/)
* [Python-docx Documentation](https://python-docx.readthedocs.io/)
* [MDN Web Docs - CSS Gradients](https://developer.mozilla.org/en-US/docs/Web/CSS/gradient)
* [Stack Overflow - Django Questions](https://stackoverflow.com/questions/tagged/django)
* [Real Python - Django Tutorials](https://realpython.com/tutorials/django/)
* [pdfplumber Documentation](https://github.com/jsvine/pdfplumber)

# Future Work
Add resume template options (modern, creative, minimal layouts)
Implement user authentication and resume storage for returning users
Integrate with ATS (Applicant Tracking System) compatibility checker
Add real-time resume preview updates as users type
Create resume version history and comparison tool
Add industry-specific resume scoring (tech vs. business vs. creative)
Implement email delivery of score reports and DOCX files
Add LinkedIn profile auto-fill capability
Create admin dashboard to track usage metrics and popular keywords
Optimize PDF export with better spacing control for single-page fitting
Add dark mode theme option
Integrate spelling/grammar checker in real-time