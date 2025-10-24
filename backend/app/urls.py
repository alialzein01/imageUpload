from django.urls import path, include
from rest_framework.routers import DefaultRouter
from rest_framework_simplejwt.views import TokenRefreshView
from . import views

# Create router for ViewSets
router = DefaultRouter()
router.register(r'classes', views.ClassViewSet)
router.register(r'students', views.StudentViewSet)
router.register(r'questions', views.QuestionViewSet, basename='question')
router.register(r'answers', views.AnswerViewSet, basename='answer')

urlpatterns = [
    path('health/', views.health_check, name='health_check'),
    path('auth/login/', views.CustomTokenObtainPairView.as_view(), name='token_obtain_pair'),
    path('auth/refresh/', TokenRefreshView.as_view(), name='token_refresh'),
    path('student/classes/', views.student_classes, name='student_classes'),
    path('student/questions/', views.student_questions, name='student_questions'),
    path('student/answers/', views.student_answers, name='student_answers'),
    path('student/submit-answer/', views.student_submit_answer, name='student_submit_answer'),
    path('instructor/class-answers/', views.instructor_class_answers, name='instructor_class_answers'),
    path('', include(router.urls)),
]
