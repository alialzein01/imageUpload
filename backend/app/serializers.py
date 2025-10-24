from rest_framework import serializers
from rest_framework.exceptions import ValidationError
from .models import Instructor, Class, Student, Question, Answer
import os


def validate_image_file(image):
    """
    Validate image file size and extension.
    Max size: 8MB
    Allowed extensions: .jpg, .jpeg, .png, .gif
    """
    if not image:
        return image
    
    # Check file size (8MB = 8 * 1024 * 1024 bytes)
    max_size = 8 * 1024 * 1024
    if image.size > max_size:
        raise ValidationError(f"File size exceeds 8MB limit. Current size: {image.size / (1024 * 1024):.2f}MB")
    
    # Check file extension
    allowed_extensions = ['.jpg', '.jpeg', '.png', '.gif']
    file_extension = os.path.splitext(image.name)[1].lower()
    
    if file_extension not in allowed_extensions:
        raise ValidationError(f"Invalid file extension. Allowed: {', '.join(allowed_extensions)}")
    
    return image


class InstructorSerializer(serializers.ModelSerializer):
    password = serializers.CharField(write_only=True, required=False)
    
    class Meta:
        model = Instructor
        fields = ['id', 'username', 'email', 'name', 'password']
    
    def create(self, validated_data):
        password = validated_data.pop('password', None)
        user = Instructor.objects.create(**validated_data)
        if password:
            user.set_password(password)
            user.save()
        return user


class ClassSerializer(serializers.ModelSerializer):
    instructor_name = serializers.SerializerMethodField()
    
    class Meta:
        model = Class
        fields = ['id', 'class_name', 'instructor', 'instructor_name', 'created_at']
        read_only_fields = ['instructor']
    
    def get_instructor_name(self, obj):
        return obj.instructor.name if obj.instructor else None
    
    def create(self, validated_data):
        # Auto-assign instructor from context
        request = self.context.get('request')
        if request and hasattr(request, 'user'):
            validated_data['instructor'] = request.user
        return super().create(validated_data)


class StudentSerializer(serializers.ModelSerializer):
    class_name = serializers.SerializerMethodField()
    
    class Meta:
        model = Student
        fields = ['id', 'name', 'phone', 'class_enrolled', 'class_name', 'created_at']
    
    def get_class_name(self, obj):
        return obj.class_enrolled.class_name if obj.class_enrolled else None


class QuestionSerializer(serializers.ModelSerializer):
    image_url = serializers.SerializerMethodField()
    class_id = serializers.IntegerField(write_only=True, required=False)
    
    class Meta:
        model = Question
        fields = ['id', 'question_text', 'image', 'image_url', 'class_related', 'class_id', 'created_at']
        read_only_fields = ['class_related']
    
    def get_image_url(self, obj):
        if obj.image:
            request = self.context.get('request')
            if request:
                return request.build_absolute_uri(obj.image.url)
        return None
    
    def validate_image(self, image):
        return validate_image_file(image)
    
    def create(self, validated_data):
        # Auto-assign class_related from context
        class_id = validated_data.pop('class_id', None)
        if class_id:
            validated_data['class_related_id'] = class_id
        return super().create(validated_data)


class AnswerSerializer(serializers.ModelSerializer):
    image_url = serializers.SerializerMethodField()
    student_name = serializers.SerializerMethodField()
    question_text = serializers.SerializerMethodField()
    student_id = serializers.IntegerField(write_only=True, required=False)
    question_id = serializers.IntegerField(write_only=True, required=False)
    
    class Meta:
        model = Answer
        fields = ['id', 'student', 'student_name', 'question', 'question_text', 'image', 'image_url', 'liked', 'student_id', 'question_id', 'created_at']
        read_only_fields = ['student', 'question']
    
    def get_image_url(self, obj):
        if obj.image:
            request = self.context.get('request')
            if request:
                return request.build_absolute_uri(obj.image.url)
        return None
    
    def get_student_name(self, obj):
        return obj.student.name if obj.student else None
    
    def get_question_text(self, obj):
        if obj.question and obj.question.question_text:
            return obj.question.question_text[:50] + ('...' if len(obj.question.question_text) > 50 else '')
        return None
    
    def validate_image(self, image):
        return validate_image_file(image)
    
    def create(self, validated_data):
        # Auto-assign student and question from context
        student_id = validated_data.pop('student_id', None)
        question_id = validated_data.pop('question_id', None)
        
        if student_id:
            validated_data['student_id'] = student_id
        if question_id:
            validated_data['question_id'] = question_id
            
        return super().create(validated_data)
