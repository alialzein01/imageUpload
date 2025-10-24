from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import AllowAny, IsAuthenticated
from rest_framework.response import Response
from rest_framework import viewsets
from rest_framework.parsers import MultiPartParser, FormParser, JSONParser
from rest_framework.exceptions import ValidationError, PermissionDenied
from rest_framework_simplejwt.views import TokenObtainPairView
from rest_framework_simplejwt.serializers import TokenObtainPairSerializer
from .models import Class, Student, Question, Answer, Instructor
from .serializers import ClassSerializer, StudentSerializer, QuestionSerializer, AnswerSerializer


class CustomTokenObtainPairSerializer(TokenObtainPairSerializer):
    @classmethod
    def get_token(cls, user):
        token = super().get_token(user)
        
        # Determine user type based on username
        user_type = 'student' if user.username.lower().startswith('student') else 'instructor'
        
        # Add custom claims
        token['user_type'] = user_type
        token['user_id'] = user.id
        token['username'] = user.username
        token['name'] = user.name
        
        return token

class CustomTokenObtainPairView(TokenObtainPairView):
    serializer_class = CustomTokenObtainPairSerializer

@api_view(['GET'])
@permission_classes([AllowAny])
def health_check(request):
    """
    Health check endpoint that returns the status of the API.
    """
    return Response({"status": "ok"})

@api_view(['GET'])
@permission_classes([IsAuthenticated])
def student_classes(request):
    """
    Get classes that the current student is enrolled in.
    """
    try:
        # Find the student record for the current user
        student = Student.objects.get(name=request.user.name)
        classes = Class.objects.filter(students=student)
        
        # Serialize the classes
        from .serializers import ClassSerializer
        serializer = ClassSerializer(classes, many=True, context={'request': request})
        
        return Response(serializer.data)
    except Student.DoesNotExist:
        return Response({"error": "Student record not found"}, status=404)
    except Exception as e:
        return Response({"error": str(e)}, status=500)

@api_view(['GET'])
@permission_classes([IsAuthenticated])
def student_questions(request):
    """
    Get questions for classes that the current student is enrolled in.
    """
    try:
        # Find the student record for the current user
        student = Student.objects.get(name=request.user.name)
        
        # Get the class_id parameter
        class_id = request.query_params.get('class_id')
        if not class_id:
            return Response({"error": "class_id parameter is required"}, status=400)
        
        # Verify the student is enrolled in this class
        try:
            class_obj = Class.objects.get(id=class_id, students=student)
        except Class.DoesNotExist:
            return Response({"error": "Class not found or you are not enrolled in this class"}, status=404)
        
        # Get questions for this class
        questions = Question.objects.filter(class_related=class_obj)
        
        # Serialize the questions
        from .serializers import QuestionSerializer
        serializer = QuestionSerializer(questions, many=True, context={'request': request})
        
        return Response(serializer.data)
    except Student.DoesNotExist:
        return Response({"error": "Student record not found"}, status=404)
    except Exception as e:
        return Response({"error": str(e)}, status=500)

@api_view(['GET'])
@permission_classes([IsAuthenticated])
def student_answers(request):
    """
    Get answers submitted by the current student for a specific class.
    """
    try:
        # Find the student record for the current user
        student = Student.objects.get(name=request.user.name)
        
        # Get the class_id parameter
        class_id = request.query_params.get('class_id')
        if not class_id:
            return Response({"error": "class_id parameter is required"}, status=400)
        
        # Verify the student is enrolled in this class
        try:
            class_obj = Class.objects.get(id=class_id, students=student)
        except Class.DoesNotExist:
            return Response({"error": "Class not found or you are not enrolled in this class"}, status=404)
        
        # Get answers submitted by this student for questions in this class
        answers = Answer.objects.filter(
            student=student,
            question__class_related=class_obj
        ).select_related('question')
        
        # Serialize the answers
        from .serializers import AnswerSerializer
        serializer = AnswerSerializer(answers, many=True, context={'request': request})
        
        return Response(serializer.data)
    except Student.DoesNotExist:
        return Response({"error": "Student record not found"}, status=404)
    except Exception as e:
        return Response({"error": str(e)}, status=500)

@api_view(['POST'])
@permission_classes([IsAuthenticated])
def student_submit_answer(request):
    """
    Submit an answer for a question by a student.
    """
    try:
        # Find the student record for the current user
        student = Student.objects.get(name=request.user.name)
        
        # Get the question_id from the request
        question_id = request.data.get('question_id')
        if not question_id:
            return Response({"error": "question_id is required"}, status=400)
        
        # Verify the question exists and the student is enrolled in the class
        try:
            question = Question.objects.select_related('class_related').get(id=question_id)
            # Check if student is enrolled in the class
            if not question.class_related.students.filter(id=student.id).exists():
                return Response({"error": "You are not enrolled in this class"}, status=403)
        except Question.DoesNotExist:
            return Response({"error": "Question not found"}, status=404)
        
        # Create the answer
        from .serializers import AnswerSerializer
        serializer = AnswerSerializer(data=request.data, context={'request': request})
        
        if serializer.is_valid():
            # Set the student and question
            serializer.save(student=student, question=question)
            return Response(serializer.data, status=201)
        else:
            return Response(serializer.errors, status=400)
            
    except Student.DoesNotExist:
        return Response({"error": "Student record not found"}, status=404)
    except Exception as e:
        return Response({"error": str(e)}, status=500)

@api_view(['GET'])
@permission_classes([IsAuthenticated])
def instructor_class_answers(request):
    """
    Get all answers for questions in a class (for instructor review).
    """
    try:
        # Get the class_id parameter
        class_id = request.query_params.get('class_id')
        if not class_id:
            return Response({"error": "class_id parameter is required"}, status=400)
        
        # Verify the instructor owns this class
        try:
            class_obj = Class.objects.get(id=class_id, instructor=request.user)
        except Class.DoesNotExist:
            return Response({"error": "Class not found or you don't have permission to view it"}, status=404)
        
        # Get all answers for questions in this class
        answers = Answer.objects.filter(
            question__class_related=class_obj
        ).select_related('student', 'question')
        
        # Serialize the answers
        from .serializers import AnswerSerializer
        serializer = AnswerSerializer(answers, many=True, context={'request': request})
        
        return Response(serializer.data)
    except Exception as e:
        return Response({"error": str(e)}, status=500)


class ClassViewSet(viewsets.ModelViewSet):
    """
    ViewSet for managing classes. Instructors can only see their own classes.
    """
    queryset = Class.objects.all()
    serializer_class = ClassSerializer
    permission_classes = [IsAuthenticated]
    
    def get_queryset(self):
        """Return only classes owned by the current instructor."""
        return Class.objects.filter(instructor=self.request.user)
    
    def perform_create(self, serializer):
        """Automatically assign the current instructor to the class."""
        serializer.save(instructor=self.request.user)


class QuestionViewSet(viewsets.ModelViewSet):
    """
    ViewSet for managing questions. Instructors can only see questions from their own classes.
    """
    queryset = Question.objects.all()
    serializer_class = QuestionSerializer
    permission_classes = [IsAuthenticated]
    parser_classes = [MultiPartParser, FormParser]
    
    def get_queryset(self):
        """Return only questions from classes owned by the current instructor."""
        queryset = Question.objects.select_related('class_related')
        
        # Filter by class_id if provided
        class_id = self.request.query_params.get('class_id')
        if class_id:
            queryset = queryset.filter(class_related_id=class_id)
        
        # Only return questions from classes owned by the current instructor
        return queryset.filter(class_related__instructor=self.request.user)
    
    def perform_create(self, serializer):
        """Verify class ownership and assign the class to the question."""
        class_id = self.request.data.get('class_id')
        if not class_id:
            raise ValidationError("class_id is required")
        
        try:
            class_obj = Class.objects.get(id=class_id, instructor=self.request.user)
        except Class.DoesNotExist:
            raise PermissionDenied("You can only create questions for your own classes")
        
        serializer.save(class_related=class_obj)


class AnswerViewSet(viewsets.ModelViewSet):
    """
    ViewSet for managing answers. Instructors can only see answers for questions in their classes.
    """
    queryset = Answer.objects.all()
    serializer_class = AnswerSerializer
    permission_classes = [IsAuthenticated]
    parser_classes = [MultiPartParser, FormParser, JSONParser]
    
    def get_queryset(self):
        """Return only answers for questions in classes owned by the current instructor."""
        # For list view, require question_id parameter
        if self.action == 'list':
            question_id = self.request.query_params.get('question_id')
            if not question_id:
                raise ValidationError("question_id parameter is required")
            
            # Verify the question belongs to a class owned by the instructor
            try:
                question = Question.objects.select_related('class_related').get(
                    id=question_id, 
                    class_related__instructor=self.request.user
                )
            except Question.DoesNotExist:
                raise PermissionDenied("Question not found or you don't have permission to view it")
            
            return Answer.objects.select_related('student', 'question').filter(question=question)
        
        # For other actions (retrieve, update, destroy), filter by instructor's classes
        return Answer.objects.select_related('student', 'question', 'question__class_related').filter(
            question__class_related__instructor=self.request.user
        )
    
    def perform_create(self, serializer):
        """Verify question and student belong to the same class."""
        question_id = self.request.data.get('question_id')
        student_id = self.request.data.get('student_id')
        
        if not question_id or not student_id:
            raise ValidationError("question_id and student_id are required")
        
        try:
            question = Question.objects.select_related('class_related').get(
                id=question_id, 
                class_related__instructor=self.request.user
            )
        except Question.DoesNotExist:
            raise PermissionDenied("Question not found or you don't have permission to create answers for it")
        
        try:
            student = Student.objects.select_related('class_enrolled').get(id=student_id)
        except Student.DoesNotExist:
            raise ValidationError("Student not found")
        
        # Verify student belongs to the same class as the question
        if student.class_enrolled != question.class_related:
            raise PermissionDenied("Student must be enrolled in the same class as the question")
        
        serializer.save(question=question, student=student)
    
    def update(self, request, *args, **kwargs):
        """Allow updating only the 'liked' field."""
        instance = self.get_object()
        
        # Only allow updating the 'liked' field
        if 'liked' not in request.data or len(request.data) > 1:
            raise ValidationError("Only the 'liked' field can be updated")
        
        return super().update(request, *args, **kwargs)
    
    def destroy(self, request, *args, **kwargs):
        """Delete the answer and its associated image file."""
        instance = self.get_object()
        
        # Delete the image file if it exists
        if instance.image:
            try:
                instance.image.delete(save=False)
            except:
                pass  # Continue even if file deletion fails
        
        return super().destroy(request, *args, **kwargs)


class StudentViewSet(viewsets.ModelViewSet):
    """
    ViewSet for managing students. Only shows students from classes owned by the current instructor.
    """
    queryset = Student.objects.all()
    serializer_class = StudentSerializer
    permission_classes = [IsAuthenticated]
    
    def get_queryset(self):
        """Return only students from classes owned by the current instructor."""
        queryset = Student.objects.filter(class_enrolled__instructor=self.request.user)
        
        # Filter by class_id if provided
        class_id = self.request.query_params.get('class_id', None)
        if class_id is not None:
            # Additional security: ensure the class belongs to the current instructor
            if Class.objects.filter(id=class_id, instructor=self.request.user).exists():
                queryset = queryset.filter(class_enrolled_id=class_id)
            else:
                # Return empty queryset if class doesn't belong to instructor
                queryset = Student.objects.none()
        
        return queryset
