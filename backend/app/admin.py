from django.contrib import admin
from django.contrib.auth.admin import UserAdmin
from .models import Instructor, Class, Student, Question, Answer


@admin.register(Instructor)
class InstructorAdmin(UserAdmin):
    """Admin configuration for Instructor model."""
    fieldsets = (
        (None, {'fields': ('username', 'password')}),
        ('Personal info', {'fields': ('first_name', 'last_name', 'name', 'email')}),
        ('Permissions', {'fields': ('is_active', 'is_staff', 'is_superuser', 'groups', 'user_permissions')}),
        ('Important dates', {'fields': ('last_login', 'date_joined')}),
    )
    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': ('username', 'name', 'email', 'password1', 'password2'),
        }),
    )


@admin.register(Class)
class ClassAdmin(admin.ModelAdmin):
    """Admin configuration for Class model."""
    list_display = ['class_name', 'instructor', 'created_at']
    list_filter = ['instructor', 'created_at']
    search_fields = ['class_name', 'instructor__name']


@admin.register(Student)
class StudentAdmin(admin.ModelAdmin):
    """Admin configuration for Student model."""
    list_display = ['name', 'phone', 'class_enrolled', 'created_at']
    list_filter = ['class_enrolled', 'created_at']
    search_fields = ['name', 'phone', 'class_enrolled__class_name']


@admin.register(Question)
class QuestionAdmin(admin.ModelAdmin):
    """Admin configuration for Question model."""
    list_display = ['question_text_short', 'class_related', 'has_image', 'created_at']
    list_filter = ['class_related', 'created_at']
    search_fields = ['question_text', 'class_related__class_name']
    
    def question_text_short(self, obj):
        """Return truncated question text for display."""
        return obj.question_text[:50] + '...' if len(obj.question_text) > 50 else obj.question_text
    question_text_short.short_description = 'Question Text'
    
    def has_image(self, obj):
        """Return whether the question has an image."""
        return bool(obj.image)
    has_image.boolean = True
    has_image.short_description = 'Has Image'


@admin.register(Answer)
class AnswerAdmin(admin.ModelAdmin):
    """Admin configuration for Answer model."""
    list_display = ['student', 'question_short', 'liked', 'created_at']
    list_filter = ['liked', 'created_at', 'student__class_enrolled']
    search_fields = ['student__name', 'question__question_text']
    
    def question_short(self, obj):
        """Return truncated question text for display."""
        return obj.question.question_text[:30] + '...' if len(obj.question.question_text) > 30 else obj.question.question_text
    question_short.short_description = 'Question'
