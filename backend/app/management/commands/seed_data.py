from django.core.management.base import BaseCommand
from app.models import Instructor, Class, Student


class Command(BaseCommand):
    help = 'Seed the database with initial data for testing'

    def handle(self, *args, **options):
        self.stdout.write('Starting data seeding...')
        
        # Clear existing data (optional - comment out if you want to keep existing data)
        self.stdout.write('Clearing existing data...')
        Student.objects.all().delete()
        Class.objects.all().delete()
        Instructor.objects.filter(username='instructor1').delete()
        
        # Create Instructor
        self.stdout.write('Creating instructor...')
        instructor, created = Instructor.objects.get_or_create(
            username='instructor1',
            defaults={
                'name': 'John Doe',
                'email': 'john@example.com',
                'is_staff': True,
                'is_superuser': True
            }
        )
        if created:
            instructor.set_password('password123')
            instructor.save()
            self.stdout.write(
                self.style.SUCCESS(f'Created instructor: {instructor.name} ({instructor.username})')
            )
        else:
            self.stdout.write(
                self.style.WARNING(f'Instructor already exists: {instructor.name} ({instructor.username})')
            )
        
        # Create Class
        self.stdout.write('Creating class...')
        math_class, created = Class.objects.get_or_create(
            class_name='Math 101',
            instructor=instructor,
            defaults={}
        )
        if created:
            self.stdout.write(
                self.style.SUCCESS(f'Created class: {math_class.class_name}')
            )
        else:
            self.stdout.write(
                self.style.WARNING(f'Class already exists: {math_class.class_name}')
            )
        
        # Create Students
        students_data = [
            {'name': 'Alice Smith', 'phone': '123-456-7890'},
            {'name': 'Bob Johnson', 'phone': '123-456-7891'},
            {'name': 'Carol White', 'phone': '123-456-7892'},
        ]
        
        self.stdout.write('Creating students...')
        for student_data in students_data:
            student, created = Student.objects.get_or_create(
                name=student_data['name'],
                class_enrolled=math_class,
                defaults={'phone': student_data['phone']}
            )
            if created:
                self.stdout.write(
                    self.style.SUCCESS(f'Created student: {student.name}')
                )
            else:
                self.stdout.write(
                    self.style.WARNING(f'Student already exists: {student.name}')
                )
        
        self.stdout.write(
            self.style.SUCCESS('\nData seeding completed successfully!')
        )
        self.stdout.write('You can now login with:')
        self.stdout.write('Username: instructor1')
        self.stdout.write('Password: password123')
