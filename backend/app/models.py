from django.contrib.auth.models import AbstractUser
from django.db import models


class Instructor(AbstractUser):
    """Instructor model extending Django's AbstractUser."""
    name = models.CharField(max_length=255)
    email = models.EmailField(unique=True)
    
    def __str__(self):
        return self.username


class Class(models.Model):
    """Class model representing a course or class."""
    class_name = models.CharField(max_length=255)
    instructor = models.ForeignKey(
        Instructor, 
        on_delete=models.CASCADE, 
        related_name='classes'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return self.class_name


class Student(models.Model):
    """Student model representing enrolled students."""
    name = models.CharField(max_length=255)
    phone = models.CharField(max_length=20, blank=True, null=True)
    class_enrolled = models.ForeignKey(
        Class, 
        on_delete=models.CASCADE, 
        related_name='students'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return self.name


class Question(models.Model):
    """Question model for class questions with optional images."""
    question_text = models.TextField()
    image = models.ImageField(upload_to='questions/', blank=True, null=True)
    class_related = models.ForeignKey(
        Class, 
        on_delete=models.CASCADE, 
        related_name='questions'
    )
    created_at = models.DateTimeField(auto_now_add=True)
    
    def __str__(self):
        return self.question_text[:50]


class Answer(models.Model):
    """Answer model for student responses to questions."""
    student = models.ForeignKey(
        Student, 
        on_delete=models.CASCADE, 
        related_name='answers'
    )
    question = models.ForeignKey(
        Question, 
        on_delete=models.CASCADE, 
        related_name='answers'
    )
    image = models.ImageField(upload_to='answers/')
    liked = models.BooleanField(default=False)
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        ordering = ['-created_at']
    
    def __str__(self):
        return f"{self.student.name} - {self.question.question_text[:30]}"
