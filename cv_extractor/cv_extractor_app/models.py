from django.db import models

# Create your models here.

class extractor(models.Model):
    uploaded_file = models.FileField(upload_to='resumes/')
