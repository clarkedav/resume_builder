from django import forms
from django.forms import ModelForm, inlineformset_factory
from .models import Education, WorkExperience, Resume


class ResumeForm(forms.ModelForm):
    """Main resume form."""
    class Meta:
        model = Resume
        exclude = ['created_at', 'updated_at']
        widgets = {
            'target_position': forms.TextInput(attrs={
                'placeholder': 'e.g. Software Developer, Marketing Manager, Data Analyst'
            }),
            'summary': forms.Textarea(attrs={
                'rows': 4,
                'placeholder': 'Write 3-4 sentences highlighting your best skills and a concrete example of what you have accomplished...'
            }),
            'skills': forms.Textarea(attrs={
                'rows': 3,
                'placeholder': 'Python, Django, HTML, CSS, JavaScript, Git, SQL, Communication, Leadership...'
            }),
        }


class EducationForm(ModelForm):
    class Meta:
        model = Education
        fields = ['degree', 'institution', 'grad_year', 'gpa', 'accomplishments']
        labels = {
            'degree': 'Degree',
            'institution': 'Institution',
            'grad_year': 'Graduation Year',
            'gpa': 'GPA (optional)',
            'accomplishments': 'Accomplishments (e.g., Dean\'s List, Awards, Key Skills Learned)',
        }


class WorkExperienceForm(forms.ModelForm):
    """Single work experience entry form."""
    class Meta:
        model = WorkExperience
        exclude = ['resume']
        widgets = {
            'description': forms.Textarea(attrs={
                'rows': 4,
                'placeholder': 'Describe your tasks and accomplishments. Use action verbs and include numbers...'
            }),
            'duration': forms.TextInput(attrs={'placeholder': 'Jan 2025 – Present'}),
            'work_location': forms.TextInput(attrs={'placeholder': 'Lagos, Nigeria'}),
        }


class UploadResumeForm(forms.Form):
    """Form for uploading an existing resume for scanning."""
    target_position = forms.CharField(
        max_length=100,
        required=False,
        widget=forms.TextInput(attrs={
            'placeholder': 'e.g. Software Developer (optional)'
        })
    )
    resume_file = forms.FileField(
        widget=forms.FileInput(attrs={'accept': '.pdf,.docx'})
    )


EducationFormSet = inlineformset_factory(
    Resume, Education,
    form=EducationForm,
    extra=2,
    can_delete=True,
    min_num=1,
    validate_min=True,
)

WorkExperienceFormSet = inlineformset_factory(
    Resume, WorkExperience,
    form=WorkExperienceForm,
    extra=2,
    can_delete=True,
)