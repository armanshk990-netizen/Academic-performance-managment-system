CREATE DATABASE IF NOT EXISTS student_marks_db;
USE student_marks_db;
CREATE TABLE IF NOT EXISTS users (user_id INT AUTO_INCREMENT PRIMARY KEY,full_name VARCHAR(120) NOT NULL,username VARCHAR(80) NOT NULL UNIQUE,password_hash VARCHAR(255) NOT NULL,role ENUM('admin','teacher','staff','student') NOT NULL DEFAULT 'teacher',is_active BOOLEAN NOT NULL DEFAULT TRUE,student_id INT NULL,password_hint VARCHAR(255) NULL,login_failures INT NOT NULL DEFAULT 0,last_failed_at DATETIME NULL,created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS settings(setting_key VARCHAR(80) PRIMARY KEY,setting_value VARCHAR(255) NOT NULL);
CREATE TABLE IF NOT EXISTS students(student_id INT AUTO_INCREMENT PRIMARY KEY,roll_no VARCHAR(30) NOT NULL UNIQUE,enrollment_no VARCHAR(50) NULL UNIQUE,student_name VARCHAR(120) NOT NULL,student_class VARCHAR(50) NOT NULL,teacher_remark VARCHAR(500) NULL,created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP);
CREATE TABLE IF NOT EXISTS subjects(subject_id INT AUTO_INCREMENT PRIMARY KEY,subject_name VARCHAR(100) NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS student_subjects(student_subject_id INT AUTO_INCREMENT PRIMARY KEY,student_id INT NOT NULL,subject_id INT NOT NULL,assigned_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,UNIQUE KEY uq_student_subject(student_id,subject_id),FOREIGN KEY(student_id) REFERENCES students(student_id) ON DELETE CASCADE,FOREIGN KEY(subject_id) REFERENCES subjects(subject_id) ON DELETE RESTRICT);
CREATE TABLE IF NOT EXISTS semesters(semester_id INT AUTO_INCREMENT PRIMARY KEY,semester_name VARCHAR(80) NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS exam_types(exam_id INT AUTO_INCREMENT PRIMARY KEY,exam_name VARCHAR(100) NOT NULL UNIQUE);
CREATE TABLE IF NOT EXISTS marks(mark_id INT AUTO_INCREMENT PRIMARY KEY,student_subject_id INT NOT NULL,semester_id INT NOT NULL,exam_id INT NOT NULL,obtained_marks DECIMAL(8,2) NOT NULL,maximum_marks DECIMAL(8,2) NOT NULL,created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,UNIQUE KEY uq_mark_record(student_subject_id,semester_id,exam_id),FOREIGN KEY(student_subject_id) REFERENCES student_subjects(student_subject_id) ON DELETE CASCADE,FOREIGN KEY(semester_id) REFERENCES semesters(semester_id) ON DELETE RESTRICT,FOREIGN KEY(exam_id) REFERENCES exam_types(exam_id) ON DELETE RESTRICT);
CREATE TABLE IF NOT EXISTS password_reset_requests(request_id INT AUTO_INCREMENT PRIMARY KEY,user_id INT NOT NULL,requested_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,status ENUM('pending','completed','dismissed') NOT NULL DEFAULT 'pending',FOREIGN KEY(user_id) REFERENCES users(user_id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS teacher_classes (user_id INT NOT NULL,class_name VARCHAR(50) NOT NULL,PRIMARY KEY(user_id,class_name),FOREIGN KEY(user_id) REFERENCES users(user_id) ON DELETE CASCADE);
CREATE TABLE IF NOT EXISTS attendance_months (attendance_month_id INT AUTO_INCREMENT PRIMARY KEY,month_key VARCHAR(30) NOT NULL UNIQUE,month_label VARCHAR(60) NOT NULL);
CREATE TABLE IF NOT EXISTS attendance (attendance_id INT AUTO_INCREMENT PRIMARY KEY,student_id INT NOT NULL,attendance_month_id INT NOT NULL,subject_id INT NULL,working_days DECIMAL(8,2) NOT NULL DEFAULT 0,present_days DECIMAL(8,2) NOT NULL DEFAULT 0,notes VARCHAR(255) NULL,UNIQUE KEY uq_attendance(student_id,attendance_month_id,subject_id),FOREIGN KEY(student_id) REFERENCES students(student_id) ON DELETE CASCADE,FOREIGN KEY(attendance_month_id) REFERENCES attendance_months(attendance_month_id) ON DELETE CASCADE,FOREIGN KEY(subject_id) REFERENCES subjects(subject_id) ON DELETE SET NULL);
INSERT IGNORE INTO settings VALUES('institution_type','college'),('institution_name','Academic Marks Management System'),('use_enrollment_no','0'),('detention_threshold','75');
-- No default/sample subjects, tests or student records are inserted.

CREATE INDEX idx_students_class ON students(student_class);
CREATE INDEX idx_marks_sem_exam ON marks(semester_id,exam_id);
CREATE INDEX idx_ss_student_subject ON student_subjects(student_id,subject_id);
CREATE INDEX idx_attendance_student_month ON attendance(student_id,attendance_month_id);
CREATE INDEX idx_reset_status_user ON password_reset_requests(user_id,status);
