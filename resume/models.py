from django.db import models


class Resume(models.Model):
    """Stores personal info, target position, summary, and skills."""
    full_name = models.CharField(max_length=100)
    email = models.EmailField()
    phone = models.CharField(max_length=20)
    location = models.CharField(max_length=100)
    linkedin = models.URLField(blank=True)
    target_position = models.CharField(max_length=100, blank=True)
    summary = models.TextField(blank=True)
    skills = models.TextField(blank=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return self.full_name

class Education(models.Model):
    resume = models.ForeignKey(Resume, on_delete=models.CASCADE, related_name='educations')
    degree = models.CharField(max_length=100)
    institution = models.CharField(max_length=200)
    grad_year = models.CharField(max_length=10)
    gpa = models.CharField(max_length=10, blank=True, null=True)
    accomplishments = models.TextField(blank=True, null=True, help_text="Key achievements, skills learned, or notable accomplishments (comma-separated)")

    def __str__(self):
        return f"{self.degree} from {self.institution}"

class WorkExperience(models.Model):
    """One work experience entry linked to a Resume."""
    resume = models.ForeignKey(Resume, on_delete=models.CASCADE, related_name='experiences')
    job_title = models.CharField(max_length=100)
    company = models.CharField(max_length=100)
    work_location = models.CharField(max_length=100, blank=True)
    duration = models.CharField(max_length=50, blank=True)
    description = models.TextField(blank=True)

    def __str__(self):
        return f"{self.job_title} at {self.company}"