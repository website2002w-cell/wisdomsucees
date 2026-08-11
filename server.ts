import express from 'express';
import path from 'path';
import fs from 'node:fs';
import crypto from 'node:crypto';
import * as XLSX from 'xlsx';
import { DatabaseSync } from 'node:sqlite';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import { GoogleGenAI } from '@google/genai';
import dotenv from 'dotenv';

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = Number(process.env.PORT) || 3000;
const DATA_DIR = path.join(process.cwd(), 'data');
const SQLITE_DB_PATH = path.join(DATA_DIR, 'school.db');
const LEGACY_JSON_PATH = path.join(DATA_DIR, 'school-db.json');

function defaultDatabaseSnapshot() {
  return {
    admissions: [
      {
        id: 'ADM-ESSUR-2026-101',
        studentName: 'K. Ananya',
        dob: '2021-04-12',
        gender: 'Female',
        grade: 'LKG',
        fatherName: 'Karthik Raja',
        motherName: 'Deepa K',
        phone: '9876543210',
        email: 'karthik@example.com',
        address: 'Main Road, Essur - 603301',
        submittedAt: new Date(Date.now() - 86400000 * 2).toISOString(),
        status: 'Interview Scheduled',
      },
    ],
    receipts: [
      {
        receiptNo: 'WNS-RCP-2026-8801',
        studentName: 'R. Kavin',
        studentId: 'WNS-2026-01',
        grade: 'Class 3',
        term: 'Term 1 Fee',
        amount: 4500,
        paymentMethod: 'UPI (GPay)',
        transactionRef: 'UPI-382910482910',
        date: new Date(Date.now() - 86400000 * 2).toLocaleDateString('en-IN'),
        status: 'VERIFIED',
      },
      {
        receiptNo: 'WNS-RCP-2026-8802',
        studentName: 'R. Kavin',
        studentId: 'WNS-2026-01',
        grade: 'Class 3',
        term: 'Book Set & Notebooks',
        amount: 1500,
        paymentMethod: 'UPI (PhonePe)',
        transactionRef: 'UPI-948201948201',
        date: new Date(Date.now() - 86400000 * 25).toLocaleDateString('en-IN'),
        status: 'VERIFIED',
      },
      {
        receiptNo: 'WNS-RCP-2026-8803',
        studentName: 'S. Priyanka',
        studentId: 'WNS-2026-02',
        grade: 'LKG',
        term: 'Term 1 Fee',
        amount: 3500,
        paymentMethod: 'UPI (GPay)',
        transactionRef: 'UPI-552019482910',
        date: new Date(Date.now() - 86400000 * 5).toLocaleDateString('en-IN'),
        status: 'VERIFIED',
      },
      {
        receiptNo: 'WNS-RCP-2026-8804',
        studentName: 'M. Yogesh',
        studentId: 'WNS-2026-03',
        grade: 'Class 5',
        term: 'Term 1 Fee',
        amount: 5000,
        paymentMethod: 'UPI (Paytm)',
        transactionRef: 'UPI-882019481102',
        date: new Date(Date.now() - 86400000 * 12).toLocaleDateString('en-IN'),
        status: 'VERIFIED',
      },
      {
        receiptNo: 'WNS-RCP-2026-8805',
        studentName: 'M. Yogesh',
        studentId: 'WNS-2026-03',
        grade: 'Class 5',
        term: 'Uniform & Accessories',
        amount: 1800,
        paymentMethod: 'UPI (GPay)',
        transactionRef: 'UPI-771920394810',
        date: new Date(Date.now() - 86400000 * 30).toLocaleDateString('en-IN'),
        status: 'VERIFIED',
      },
      {
        receiptNo: 'WNS-RCP-2026-8806',
        studentName: 'K. Ananya',
        studentId: 'WNS-2026-04',
        grade: 'LKG',
        term: 'Term 1 Fee',
        amount: 3500,
        paymentMethod: 'UPI (BHIM)',
        transactionRef: 'UPI-481920391820',
        date: new Date(Date.now() - 86400000 * 1).toLocaleDateString('en-IN'),
        status: 'VERIFIED',
      },
    ],
    studentProfiles: {
      'WNS-2026-01': {
        id: 'WNS-2026-01',
        name: 'R. Kavin',
        grade: 'Class 3 - A',
        rollNo: '01',
        phone: '9876543210',
        dob: '2018-06-15',
        fatherName: 'R. Saravanan',
        motherName: 'S. Sundari',
        classTeacher: 'Mrs. M. Lakshmi, M.Sc., B.Ed.',
        attendancePercentage: 96.5,
        totalDays: 85,
        presentDays: 82,
        bloodGroup: 'O+',
        feeStatus: 'Paid for Term 1',
        term1Marks: [
          { subject: 'Tamil', mark: 92, max: 100, grade: 'A+' },
          { subject: 'English', mark: 88, max: 100, grade: 'A' },
          { subject: 'Mathematics', mark: 95, max: 100, grade: 'A+' },
          { subject: 'Science', mark: 90, max: 100, grade: 'A+' },
          { subject: 'Social Studies', mark: 86, max: 100, grade: 'A' },
          { subject: 'Computer Literacy', mark: 98, max: 100, grade: 'A+' },
        ],
        homework: [
          { id: 'hw-1', subject: 'Mathematics', title: 'Exercise 3.2 Multiplication tables 6 to 9', dueDate: 'Tomorrow', status: 'Pending' },
          { id: 'hw-2', subject: 'Tamil', title: 'Thirukkural memorization & write 2 times', dueDate: 'Completed', status: 'Submitted' },
          { id: 'hw-3', subject: 'English', title: 'Read Chapter 4 story & answer Q1 to Q5', dueDate: '23 Jul 2026', status: 'Pending' },
        ],
        teacherNotes: 'Kavin is bright, observant, and participates enthusiastically in class activities and math quizzes!',
      },
      'WNS-2026-02': {
        id: 'WNS-2026-02',
        name: 'S. Priyanka',
        grade: 'LKG - Blossom',
        rollNo: '14',
        phone: '9876543211',
        dob: '2021-04-10',
        fatherName: 'S. Sundaram',
        motherName: 'S. Kavitha',
        classTeacher: 'Mrs. V. Radha, B.A., D.T.Ed.',
        attendancePercentage: 98.0,
        totalDays: 85,
        presentDays: 83,
        bloodGroup: 'B+',
        feeStatus: 'Paid for Annual Term',
        term1Marks: [
          { subject: 'English Alphabets & Phonics', mark: 98, max: 100, grade: 'A+' },
          { subject: 'Tamil Rhymes & Vowels', mark: 95, max: 100, grade: 'A+' },
          { subject: 'Number Counting (1-50)', mark: 96, max: 100, grade: 'A+' },
          { subject: 'Drawing & Coloring', mark: 100, max: 100, grade: 'A+' },
          { subject: 'Good Habits & Rhymes', mark: 94, max: 100, grade: 'A+' },
        ],
        homework: [
          { id: 'hw-10', subject: 'Drawing', title: 'Color the Mango and Sun worksheet', dueDate: 'Tomorrow', status: 'Pending' },
          { id: 'hw-11', subject: 'Phonics', title: 'Practice A to Z phonetic songs', dueDate: 'Completed', status: 'Submitted' },
        ],
        teacherNotes: 'Priyanka loves storytelling and rhymes! Excellent rhythm and social interaction.',
      },
      'WNS-2026-03': {
        id: 'WNS-2026-03',
        name: 'M. Yogesh',
        grade: 'Class 5 - A',
        rollNo: '22',
        phone: '9876543212',
        dob: '2016-08-20',
        fatherName: 'M. Manikandan',
        motherName: 'M. Selvi',
        classTeacher: 'Mr. P. Ramesh, M.A., B.Ed.',
        attendancePercentage: 94.2,
        totalDays: 85,
        presentDays: 80,
        bloodGroup: 'A1+',
        feeStatus: 'Term 2 Fee Due',
        term1Marks: [
          { subject: 'Tamil', mark: 85, max: 100, grade: 'A' },
          { subject: 'English', mark: 91, max: 100, grade: 'A+' },
          { subject: 'Mathematics', mark: 89, max: 100, grade: 'A' },
          { subject: 'Science', mark: 94, max: 100, grade: 'A+' },
          { subject: 'Social Science', mark: 88, max: 100, grade: 'A' },
          { subject: 'Computer Applications', mark: 96, max: 100, grade: 'A+' },
        ],
        homework: [
          { id: 'hw-20', subject: 'Science', title: 'Draw and label parts of a Plant Cell', dueDate: '24 Jul 2026', status: 'Pending' },
          { id: 'hw-21', subject: 'Mathematics', title: 'Fractions & Decimals Worksheet 5', dueDate: 'Tomorrow', status: 'Pending' },
        ],
        teacherNotes: 'Yogesh is showing high leadership in computer lab sessions and sports activities!',
      },
    },
    feeReminders: [
      {
        id: 'REM-2026-901',
        studentId: 'WNS-2026-01',
        studentName: 'R. Kavin',
        parentPhone: '9876543210',
        term: 'Term 2 Tuition Fee',
        amount: 3500,
        channel: 'push',
        sentAt: new Date(Date.now() - 3600000 * 24 * 2).toLocaleString('en-IN'),
        status: 'DELIVERED',
        messageText: '🔔 [Push Notification] Gentle Fee Reminder: Term 2 Fee ₹3,500 for R. Kavin is due in 3 days (15-Aug-2026). Tap to pay via GPay/PhonePe.',
      },
    ],
    notificationCampaignHistory: [
      {
        id: 'CMP-2026-8801',
        title: 'Term 2 Fee Warning - August Batch',
        channel: 'WhatsApp + SMS',
        recipientCount: 6,
        totalAmountOverdue: 23400,
        messagePreview: 'Dear Parent, gentle reminder regarding pending Term 2 fees. Please pay online via GPay/PhonePe to avoid late charges.',
        sentAt: '08-Aug-2026 10:30 AM',
        sentBy: 'Admin R. Saravanan',
        status: 'Delivered (100%)',
      },
      {
        id: 'CMP-2026-8795',
        title: 'Monthly Installment #2 Payment Alert',
        channel: 'WhatsApp',
        recipientCount: 4,
        totalAmountOverdue: 14800,
        messagePreview: 'Wisdom School Notice: Monthly fee installment #2 is due. Click link to complete UPI payment.',
        sentAt: '01-Aug-2026 04:15 PM',
        sentBy: 'Admin Office Staff',
        status: 'Delivered (100%)',
      },
    ],
    feeCategories: [
      {
        id: 'fee-admission',
        key: 'admission',
        name: 'Admission Fee (New Enrollments)',
        amount: 2500,
        description: 'One-time registration, ID card, prospectus & admission processing',
        deletable: false,
        editable: true,
        categoryType: 'Admission',
        frequency: 'One-Time',
      },
      {
        id: 'fee-sports',
        key: 'sports',
        name: 'Sports & Annual Athletic Fee',
        amount: 1200,
        description: 'Annual sports equipment, physical education coaching & athletic meet',
        deletable: false,
        editable: true,
        categoryType: 'Sports',
        frequency: 'Annual',
      },
      {
        id: 'fee-van-1',
        key: 'van_route_1',
        name: 'School Van / Transport Fee (Route 1 - Essur Local)',
        amount: 1500,
        description: 'Doorstep pickup & drop-off for Essur village and surrounding areas',
        deletable: true,
        editable: true,
        categoryType: 'Van / Transport',
        frequency: 'Term-wise',
      },
      {
        id: 'fee-van-2',
        key: 'van_route_2',
        name: 'School Van / Transport Fee (Route 2 - Cheyyur / Highway)',
        amount: 1800,
        description: 'Van service covering Cheyyur highway and suburban stops',
        deletable: true,
        editable: true,
        categoryType: 'Van / Transport',
        frequency: 'Term-wise',
      },
      {
        id: 'fee-term1',
        key: 'term1',
        name: 'Term 1 Tuition & Activity Fee',
        amount: 4500,
        description: 'Academic tuition & classroom learning for Term 1',
        deletable: false,
        editable: true,
        categoryType: 'Tuition',
        frequency: 'Term-wise',
      },
      {
        id: 'fee-term2',
        key: 'term2',
        name: 'Term 2 Tuition & Activity Fee',
        amount: 4500,
        description: 'Academic tuition & classroom learning for Term 2',
        deletable: false,
        editable: true,
        categoryType: 'Tuition',
        frequency: 'Term-wise',
      },
      {
        id: 'fee-term3',
        key: 'term3',
        name: 'Term 3 Tuition & Activity Fee',
        amount: 3800,
        description: 'Academic tuition & classroom learning for Term 3',
        deletable: false,
        editable: true,
        categoryType: 'Tuition',
        frequency: 'Term-wise',
      },
      {
        id: 'fee-books',
        key: 'books',
        name: 'Book Set, Workbooks & Stationary',
        amount: 2000,
        description: 'Tamil Nadu Board textbooks, note copies and homework diaries',
        deletable: false,
        editable: true,
        categoryType: 'Books',
        frequency: 'Annual',
      },
      {
        id: 'fee-uniform',
        key: 'uniform',
        name: 'Uniform, Sports Jersey & Accessories',
        amount: 1500,
        description: '2 sets school uniform, sports jersey and school badge',
        deletable: false,
        editable: true,
        categoryType: 'Uniform',
        frequency: 'Annual',
      },
    ],
  };
}

function getDb() {
  return new DatabaseSync(SQLITE_DB_PATH);
}

function hashPassword(password: string) {
  return crypto.pbkdf2Sync(password, 'wisdom-school-salt-v1', 100000, 64, 'sha512').toString('hex');
}

function ensureDatabaseFiles() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }

  if (!fs.existsSync(LEGACY_JSON_PATH)) {
    fs.writeFileSync(LEGACY_JSON_PATH, JSON.stringify(defaultDatabaseSnapshot(), null, 2));
  }

  const db = getDb();
  db.exec(`CREATE TABLE IF NOT EXISTS app_data (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS admin_roles (
    id TEXT PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    permissions TEXT NOT NULL DEFAULT '[]'
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS admin_users (
    id TEXT PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    name TEXT NOT NULL,
    role_id TEXT NOT NULL,
    created_at TEXT NOT NULL,
    FOREIGN KEY (role_id) REFERENCES admin_roles(id)
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS admin_sessions (
    id TEXT PRIMARY KEY,
    user_id TEXT NOT NULL,
    token TEXT NOT NULL UNIQUE,
    created_at TEXT NOT NULL,
    expires_at TEXT NOT NULL,
    revoked INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (user_id) REFERENCES admin_users(id)
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS admissions (
    id TEXT PRIMARY KEY,
    student_name TEXT NOT NULL,
    dob TEXT,
    gender TEXT,
    grade TEXT NOT NULL,
    father_name TEXT,
    mother_name TEXT,
    phone TEXT,
    email TEXT,
    address TEXT,
    previous_school TEXT,
    submitted_at TEXT NOT NULL,
    status TEXT NOT NULL
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS fee_receipts (
    id TEXT PRIMARY KEY,
    receipt_no TEXT NOT NULL UNIQUE,
    student_name TEXT NOT NULL,
    student_id TEXT,
    grade TEXT,
    term TEXT,
    amount REAL NOT NULL,
    payment_method TEXT,
    transaction_ref TEXT,
    date TEXT NOT NULL,
    status TEXT NOT NULL
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS student_records (
    id TEXT PRIMARY KEY,
    student_id TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    grade TEXT,
    roll_no TEXT,
    phone TEXT,
    dob TEXT,
    father_name TEXT,
    mother_name TEXT,
    class_teacher TEXT,
    attendance_percentage REAL,
    total_days INTEGER,
    present_days INTEGER,
    blood_group TEXT,
    fee_status TEXT,
    teacher_notes TEXT,
    important_notes TEXT,
    metadata TEXT NOT NULL DEFAULT '{}'
  );`);

  db.exec(`CREATE TABLE IF NOT EXISTS fee_categories (
    id TEXT PRIMARY KEY,
    key_name TEXT NOT NULL UNIQUE,
    name TEXT NOT NULL,
    amount REAL NOT NULL,
    description TEXT,
    deletable INTEGER NOT NULL DEFAULT 0,
    editable INTEGER NOT NULL DEFAULT 1,
    category_type TEXT,
    frequency TEXT
  );`);

  const roleCount = db.prepare('SELECT COUNT(*) AS count FROM admin_roles').get() as { count: number };
  if (roleCount.count === 0) {
    db.prepare('INSERT INTO admin_roles (id, name, permissions) VALUES (?, ?, ?)').run('school_admin', 'school_admin', JSON.stringify(['read','write','delete']));
  }

  const adminCount = db.prepare('SELECT COUNT(*) AS count FROM admin_users').get() as { count: number };
  if (adminCount.count === 0) {
    const roleId = db.prepare('SELECT id FROM admin_roles WHERE name = ?').get('school_admin') as { id: string } | undefined;
    if (roleId) {
      db.prepare('INSERT INTO admin_users (id, username, password_hash, name, role_id, created_at) VALUES (?, ?, ?, ?, ?, ?)')
        .run('admin-wisdom', 'WISDOM', hashPassword('WISDOM2002'), 'R. SARAVANAN', roleId.id, new Date().toISOString());
    }
  }

  const categoryCount = db.prepare('SELECT COUNT(*) AS count FROM fee_categories').get() as { count: number };
  if (categoryCount.count === 0) {
    const defaultCategories = defaultDatabaseSnapshot().feeCategories as any[];
    const insertCategory = db.prepare('INSERT INTO fee_categories (id, key_name, name, amount, description, deletable, editable, category_type, frequency) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)');
    for (const category of defaultCategories) {
      insertCategory.run(
        category.id,
        category.key,
        category.name,
        Number(category.amount),
        category.description || '',
        category.deletable ? 1 : 0,
        category.editable ? 1 : 0,
        category.categoryType,
        category.frequency,
      );
    }
  }

  const legacyRows = fs.existsSync(LEGACY_JSON_PATH) ? JSON.parse(fs.readFileSync(LEGACY_JSON_PATH, 'utf-8')) : null;
  if (legacyRows && (db.prepare('SELECT COUNT(*) AS count FROM admissions').get() as { count: number }).count === 0) {
    const admissionInsert = db.prepare('INSERT INTO admissions (id, student_name, dob, gender, grade, father_name, mother_name, phone, email, address, previous_school, submitted_at, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)');
    for (const item of Array.isArray(legacyRows.admissions) ? legacyRows.admissions : []) {
      admissionInsert.run(item.id, item.studentName, item.dob || '', item.gender || '', item.grade || '', item.fatherName || '', item.motherName || '', item.phone || '', item.email || '', item.address || '', item.previousSchool || '', item.submittedAt || new Date().toISOString(), item.status || 'Pending Review');
    }

    const receiptInsert = db.prepare('INSERT INTO fee_receipts (id, receipt_no, student_name, student_id, grade, term, amount, payment_method, transaction_ref, date, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)');
    for (const item of Array.isArray(legacyRows.receipts) ? legacyRows.receipts : []) {
      receiptInsert.run(item.id || item.receiptNo, item.receiptNo, item.studentName, item.studentId || '', item.grade || '', item.term || '', Number(item.amount || 0), item.paymentMethod || '', item.transactionRef || '', item.date || new Date().toLocaleDateString('en-IN'), item.status || 'VERIFIED');
    }

    const studentInsert = db.prepare('INSERT INTO student_records (id, student_id, name, grade, roll_no, phone, dob, father_name, mother_name, class_teacher, attendance_percentage, total_days, present_days, blood_group, fee_status, teacher_notes, important_notes, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)');
    const studentProfilesMap = legacyRows.studentProfiles && typeof legacyRows.studentProfiles === 'object' ? legacyRows.studentProfiles : {};
    for (const student of Object.values(studentProfilesMap) as any[]) {
      studentInsert.run(
        student.id || crypto.randomUUID(),
        student.id || `STU-${Date.now()}`,
        student.name || 'Student',
        student.grade || '',
        student.rollNo || '',
        student.phone || '',
        student.dob || '',
        student.fatherName || '',
        student.motherName || '',
        student.classTeacher || '',
        Number(student.attendancePercentage || 0),
        Number(student.totalDays || 0),
        Number(student.presentDays || 0),
        student.bloodGroup || '',
        student.feeStatus || '',
        student.teacherNotes || '',
        student.importantNotes || '',
        JSON.stringify({
          term1Marks: Array.isArray(student.term1Marks) ? student.term1Marks : [],
          homework: Array.isArray(student.homework) ? student.homework : [],
          ...student,
        }),
      );
    }
  }
}

function refreshInMemoryStores() {
  const db = getDb();
  admissionsDb = db.prepare('SELECT * FROM admissions ORDER BY submitted_at DESC').all() as unknown as AdmissionApplication[];
  receiptsDb = db.prepare('SELECT * FROM fee_receipts ORDER BY date DESC').all() as unknown as FeeReceipt[];

  const rows = db.prepare('SELECT * FROM student_records ORDER BY name ASC').all() as any[];
  studentProfiles = {} as Record<string, any>;
  for (const row of rows) {
    const metadata = row.metadata ? JSON.parse(row.metadata) : {};
    studentProfiles[row.student_id] = {
      id: row.student_id,
      name: row.name,
      grade: row.grade,
      rollNo: row.roll_no,
      phone: row.phone,
      dob: row.dob,
      fatherName: row.father_name,
      motherName: row.mother_name,
      classTeacher: row.class_teacher,
      attendancePercentage: row.attendance_percentage,
      totalDays: row.total_days,
      presentDays: row.present_days,
      bloodGroup: row.blood_group,
      feeStatus: row.fee_status,
      teacherNotes: row.teacher_notes,
      importantNotes: row.important_notes,
      ...metadata,
      term1Marks: Array.isArray(metadata.term1Marks) ? metadata.term1Marks : [],
      homework: Array.isArray(metadata.homework) ? metadata.homework : [],
    };
  }

  feeCategoriesDb = db.prepare('SELECT * FROM fee_categories ORDER BY name ASC').all() as any[];
  const legacyJson = fs.readFileSync(LEGACY_JSON_PATH, 'utf-8');
  const legacyData = legacyJson ? JSON.parse(legacyJson || '{}') : {};
  feeRemindersDb = Array.isArray(legacyData.feeReminders) ? legacyData.feeReminders : [];
  notificationCampaignHistoryDb = Array.isArray(legacyData.notificationCampaignHistory) ? legacyData.notificationCampaignHistory : [];
}

function loadDatabase() {
  ensureDatabaseFiles();
  refreshInMemoryStores();
  return {
    admissions: admissionsDb,
    receipts: receiptsDb,
    studentProfiles,
    feeReminders: feeRemindersDb,
    notificationCampaignHistory: notificationCampaignHistoryDb,
    feeCategories: feeCategoriesDb,
  };
}

function persistDatabase() {
  const snapshot = {
    admissions: admissionsDb,
    receipts: receiptsDb,
    studentProfiles,
    feeReminders: feeRemindersDb,
    notificationCampaignHistory: notificationCampaignHistoryDb,
    feeCategories: feeCategoriesDb,
  };

  fs.writeFileSync(LEGACY_JSON_PATH, JSON.stringify(snapshot, null, 2));

  const db = getDb();
  const writeAdmission = db.prepare('INSERT INTO admissions (id, student_name, dob, gender, grade, father_name, mother_name, phone, email, address, previous_school, submitted_at, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET student_name=excluded.student_name, dob=excluded.dob, gender=excluded.gender, grade=excluded.grade, father_name=excluded.father_name, mother_name=excluded.mother_name, phone=excluded.phone, email=excluded.email, address=excluded.address, previous_school=excluded.previous_school, submitted_at=excluded.submitted_at, status=excluded.status');
  for (const application of admissionsDb) {
    writeAdmission.run(application.id, application.studentName, application.dob || '', application.gender || '', application.grade || '', application.fatherName || '', application.motherName || '', application.phone || '', application.email || '', application.address || '', application.previousSchool || '', application.submittedAt || new Date().toISOString(), application.status || 'Pending Review');
  }

  const writeReceipt = db.prepare('INSERT INTO fee_receipts (id, receipt_no, student_name, student_id, grade, term, amount, payment_method, transaction_ref, date, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET receipt_no=excluded.receipt_no, student_name=excluded.student_name, student_id=excluded.student_id, grade=excluded.grade, term=excluded.term, amount=excluded.amount, payment_method=excluded.payment_method, transaction_ref=excluded.transaction_ref, date=excluded.date, status=excluded.status');
  for (const receipt of receiptsDb) {
    writeReceipt.run(receipt.id || receipt.receiptNo, receipt.receiptNo, receipt.studentName, receipt.studentId || '', receipt.grade || '', receipt.term || '', Number(receipt.amount || 0), receipt.paymentMethod || '', receipt.transactionRef || '', receipt.date || new Date().toLocaleDateString('en-IN'), receipt.status || 'VERIFIED');
  }

  const studentUpsert = db.prepare('INSERT INTO student_records (id, student_id, name, grade, roll_no, phone, dob, father_name, mother_name, class_teacher, attendance_percentage, total_days, present_days, blood_group, fee_status, teacher_notes, important_notes, metadata) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(student_id) DO UPDATE SET id=excluded.id, name=excluded.name, grade=excluded.grade, roll_no=excluded.roll_no, phone=excluded.phone, dob=excluded.dob, father_name=excluded.father_name, mother_name=excluded.mother_name, class_teacher=excluded.class_teacher, attendance_percentage=excluded.attendance_percentage, total_days=excluded.total_days, present_days=excluded.present_days, blood_group=excluded.blood_group, fee_status=excluded.fee_status, teacher_notes=excluded.teacher_notes, important_notes=excluded.important_notes, metadata=excluded.metadata');
  for (const [studentId, student] of Object.entries(studentProfiles)) {
    const metadata = {
      term1Marks: Array.isArray(student.term1Marks) ? student.term1Marks : [],
      homework: Array.isArray(student.homework) ? student.homework : [],
      ...student,
    };
    delete metadata.id;
    delete metadata.name;
    delete metadata.grade;
    delete metadata.rollNo;
    delete metadata.phone;
    delete metadata.dob;
    delete metadata.fatherName;
    delete metadata.motherName;
    delete metadata.classTeacher;
    delete metadata.attendancePercentage;
    delete metadata.totalDays;
    delete metadata.presentDays;
    delete metadata.bloodGroup;
    delete metadata.feeStatus;
    delete metadata.teacherNotes;
    delete metadata.importantNotes;
    studentUpsert.run(
      student.id || studentId,
      studentId,
      student.name || 'Student',
      student.grade || '',
      student.rollNo || '',
      student.phone || '',
      student.dob || '',
      student.fatherName || '',
      student.motherName || '',
      student.classTeacher || '',
      Number(student.attendancePercentage || 0),
      Number(student.totalDays || 0),
      Number(student.presentDays || 0),
      student.bloodGroup || '',
      student.feeStatus || '',
      student.teacherNotes || '',
      student.importantNotes || '',
      JSON.stringify(metadata),
    );
  }

  const categoryUpsert = db.prepare('INSERT INTO fee_categories (id, key_name, name, amount, description, deletable, editable, category_type, frequency) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET key_name=excluded.key_name, name=excluded.name, amount=excluded.amount, description=excluded.description, deletable=excluded.deletable, editable=excluded.editable, category_type=excluded.category_type, frequency=excluded.frequency');
  for (const category of feeCategoriesDb) {
    categoryUpsert.run(category.id, category.key || category.key_name || category.id, category.name, Number(category.amount || 0), category.description || '', category.deletable ? 1 : 0, category.editable ? 1 : 0, category.categoryType || category.category_type || '', category.frequency || '');
  }

  db.prepare('DELETE FROM app_data').run();
  const appDataInsert = db.prepare('INSERT INTO app_data (key, value) VALUES (?, ?)');
  for (const [key, value] of Object.entries(snapshot)) {
    appDataInsert.run(key, JSON.stringify(value));
  }
}

let admissionsDb: AdmissionApplication[] = [];
let receiptsDb: FeeReceipt[] = [];
let studentProfiles: Record<string, any> = {};
let feeRemindersDb: FeeReminderLog[] = [];
let notificationCampaignHistoryDb: any[] = [];
let feeCategoriesDb: any[] = [];

const database = loadDatabase();
admissionsDb = database.admissions;
receiptsDb = database.receipts;
studentProfiles = database.studentProfiles as Record<string, any>;
feeRemindersDb = database.feeReminders;
notificationCampaignHistoryDb = database.notificationCampaignHistory;
feeCategoriesDb = database.feeCategories;

app.use(express.json({ limit: '2mb' }));

function getRequestToken(req: express.Request) {
  const authHeader = req.headers.authorization || '';
  const match = authHeader.match(/^Bearer\s+(.+)$/i);
  if (match) return match[1];
  return typeof req.headers['x-session-token'] === 'string' ? req.headers['x-session-token'] : '';
}

function getAdminUserByToken(token: string) {
  if (!token) return null;
  const db = getDb();
  const row = db.prepare('SELECT * FROM admin_sessions WHERE token = ? AND revoked = 0 AND expires_at > ?').get(token, new Date().toISOString()) as any;
  if (!row) return null;
  const user = db.prepare('SELECT u.*, r.name AS role_name, r.permissions AS role_permissions FROM admin_users u JOIN admin_roles r ON r.id = u.role_id WHERE u.id = ?').get(row.user_id) as any;
  if (!user) return null;

  let permissions: string[] = [];
  try {
    permissions = Array.isArray(JSON.parse(user.role_permissions || '[]')) ? JSON.parse(user.role_permissions || '[]') : [];
  } catch {
    permissions = [];
  }

  return {
    id: user.id,
    username: user.username,
    name: user.name,
    role: user.role_name,
    permissions,
  };
}

function requireAdminAuth(req: express.Request, res: express.Response, next: express.NextFunction) {
  const token = getRequestToken(req);
  const admin = getAdminUserByToken(token);
  if (!admin) {
    res.status(401).json({ success: false, message: 'Admin authentication required.' });
    return;
  }
  (req as any).adminUser = admin;
  next();
}

function normalizeExportRow(row: Record<string, any>) {
  const flat: Record<string, any> = {};
  for (const [key, value] of Object.entries(row || {})) {
    if (value === undefined || value === null) {
      flat[key] = '';
    } else if (typeof value === 'object') {
      flat[key] = JSON.stringify(value);
    } else {
      flat[key] = value;
    }
  }
  return flat;
}

function getCsvFromRecords(records: Record<string, any>[]) {
  if (!records || records.length === 0) {
    return '';
  }

  const headers = Object.keys(normalizeExportRow(records[0]));
  const rows = records.map((record) => {
    const normalized = normalizeExportRow(record);
    return headers.map((header) => {
      const value = normalized[header] ?? '';
      const stringValue = String(value).replace(/\r?\n/g, ' ').replace(/"/g, '""');
      return `"${stringValue}"`;
    }).join(',');
  });

  return [headers.join(','), ...rows].join('\n');
}

// Initialize Gemini AI Client lazily or safely
function getGeminiClient() {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    return null;
  }
  return new GoogleGenAI({
    apiKey: apiKey,
    httpOptions: {
      headers: {
        'User-Agent': 'aistudio-build',
      },
    },
  });
}

// In-memory data store for admissions and receipts (transient session/demo storage)
interface AdmissionApplication {
  id: string;
  studentName: string;
  dob: string;
  gender: string;
  grade: string;
  fatherName: string;
  motherName: string;
  phone: string;
  email: string;
  address: string;
  previousSchool?: string;
  submittedAt: string;
  status: 'Pending Review' | 'Accepted' | 'Interview Scheduled';
}

interface FeeReceipt {
  id?: string;
  receiptNo: string;
  studentName: string;
  studentId: string;
  grade: string;
  term: string;
  amount: number;
  paymentMethod: string;
  transactionRef: string;
  date: string;
  status: 'VERIFIED' | 'PENDING_VERIFICATION';
}

// API Routes

// Health Check
app.get('/api/health', (req, res) => {
  res.json({
    status: 'online',
    schoolName: 'Wisdom Nursery and Primary School',
    location: 'Essur - 603301',
    motto: 'Learn Today, Lead Tomorrow.',
    admin: 'R. SARAVANAN',
  });
});

// Student Portal Login API (By Phone & DOB OR Student ID)
app.post('/api/student/login', (req, res) => {
  const { phone, dob, studentId } = req.body || {};

  // If studentId provided
  if (studentId) {
    const cleanId = String(studentId).trim().toUpperCase();
    const found = studentProfiles[cleanId];
    if (found) {
      res.json({ success: true, student: found });
      return;
    }
  }

  // If Phone + DOB provided
  if (phone) {
    const cleanPhone = String(phone).replace(/\D/g, '').slice(-10);
    const cleanDob = String(dob || '').trim();

    const matched = Object.values(studentProfiles).find((st: any) => {
      const stPhone = String(st.phone || '').replace(/\D/g, '').slice(-10);
      const phoneMatch = stPhone === cleanPhone || stPhone.endsWith(cleanPhone) || cleanPhone.endsWith(stPhone);
      const dobMatch = !cleanDob || st.dob === cleanDob || st.dob?.replace(/-/g, '') === cleanDob.replace(/-/g, '');
      return phoneMatch && dobMatch;
    });

    if (matched) {
      res.json({ success: true, student: matched });
      return;
    }
  }

  res.status(401).json({
    success: false,
    message: 'No student found matching the provided details. Try Mobile: 9876543210 & DOB: 2018-06-15, or Student ID: WNS-2026-01',
  });
});

// Admin Login API (Password: WISDOM2002, Username: WISDOM)
app.post('/api/admin/login', (req, res) => {
  const { username, password } = req.body || {};
  const cleanUsername = String(username || '').trim().toUpperCase();
  const cleanPassword = String(password || '').trim();

  if (!cleanUsername || !cleanPassword) {
    res.status(401).json({
      success: false,
      message: 'Username and password are required.',
    });
    return;
  }

  const db = getDb();
  const user = db.prepare(`SELECT u.*, r.name AS role_name, r.permissions AS role_permissions
    FROM admin_users u
    JOIN admin_roles r ON r.id = u.role_id
    WHERE u.username = ?`).get(cleanUsername) as any;

  if (!user || user.password_hash !== hashPassword(cleanPassword)) {
    res.status(401).json({
      success: false,
      message: 'Invalid Admin Credentials! Username is "WISDOM" and Password is "WISDOM2002".',
    });
    return;
  }

  const token = crypto.randomBytes(32).toString('hex');
  const expiresAt = new Date(Date.now() + 12 * 60 * 60 * 1000).toISOString();
  db.prepare('INSERT INTO admin_sessions (id, user_id, token, created_at, expires_at, revoked) VALUES (?, ?, ?, ?, ?, 0)').run(
    crypto.randomUUID(),
    user.id,
    token,
    new Date().toISOString(),
    expiresAt,
  );

  res.json({
    success: true,
    token,
    admin: {
      username: user.username,
      name: user.name,
      role: user.role_name,
      school: 'Wisdom Nursery and Primary School, Essur - 603301',
    },
    message: 'Admin Authentication Successful!',
  });
});

// Admin Dashboard Data Fetch API
app.get('/api/admin/data', requireAdminAuth, (req, res) => {
  res.json({
    success: true,
    stats: {
      totalAdmissions: admissionsDb.length,
      totalFeeReceipts: receiptsDb.length,
      totalStudents: Object.keys(studentProfiles).length,
      schoolName: 'Wisdom Nursery and Primary School, Essur - 603301',
    },
    admissions: admissionsDb,
    receipts: receiptsDb,
    students: Object.values(studentProfiles),
  });
});

app.get('/api/admin/export/:collection', requireAdminAuth, (req, res) => {
  const collection = String(req.params.collection || '').toLowerCase();
  const format = String(req.query.format || 'csv').toLowerCase();

  let records: Record<string, any>[] = [];
  if (collection === 'admissions') {
    records = admissionsDb as any[];
  } else if (collection === 'fees' || collection === 'receipts') {
    records = receiptsDb as any[];
  } else if (collection === 'students') {
    records = Object.values(studentProfiles);
  } else {
    res.status(400).json({ success: false, message: 'Unsupported export collection. Use students, admissions, or fees.' });
    return;
  }

  if (format === 'xlsx' || format === 'excel') {
    const ws = XLSX.utils.json_to_sheet(records.map(normalizeExportRow));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, collection);
    const fileBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=wisdom_${collection}.xlsx`);
    res.send(fileBuffer);
    return;
  }

  const csv = getCsvFromRecords(records.map(normalizeExportRow));
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename=wisdom_${collection}.csv`);
  res.send(csv);
});

app.post('/api/admin/import/:collection', requireAdminAuth, (req, res) => {
  const collection = String(req.params.collection || '').toLowerCase();
  const records = Array.isArray(req.body?.records) ? req.body.records : Array.isArray(req.body) ? req.body : [];

  if (!records.length) {
    res.status(400).json({ success: false, message: 'No records provided for import.' });
    return;
  }

  try {
    if (collection === 'students') {
      for (const item of records) {
        const studentId = String(item.studentId || item.id || `WNS-${Date.now()}-${Math.random().toString().slice(2, 6)}`);
        studentProfiles[studentId] = {
          id: studentId,
          name: item.name || 'Student',
          grade: item.grade || 'Class 1',
          rollNo: item.rollNo || item.roll_no || '0',
          phone: item.phone || '9876543210',
          dob: item.dob || '2018-06-15',
          fatherName: item.fatherName || 'Parent',
          motherName: item.motherName || 'Parent',
          classTeacher: item.classTeacher || 'Class Teacher',
          attendancePercentage: Number(item.attendancePercentage || 0),
          totalDays: Number(item.totalDays || 0),
          presentDays: Number(item.presentDays || 0),
          bloodGroup: item.bloodGroup || 'O+',
          feeStatus: item.feeStatus || 'Pending',
          teacherNotes: item.teacherNotes || '',
          importantNotes: item.importantNotes || '',
          term1Marks: Array.isArray(item.term1Marks) ? item.term1Marks : [],
          homework: Array.isArray(item.homework) ? item.homework : [],
        };
      }
      persistDatabase();
      res.json({ success: true, importedCount: records.length, collection: 'students' });
      return;
    }

    if (collection === 'admissions') {
      for (const item of records) {
        const id = String(item.id || `ADM-ESSUR-${Date.now()}-${Math.random().toString().slice(2, 6)}`);
        const admission = {
          id,
          studentName: item.studentName || 'Student',
          dob: item.dob || 'N/A',
          gender: item.gender || 'Unspecified',
          grade: item.grade || 'LKG',
          fatherName: item.fatherName || 'Parent',
          motherName: item.motherName || 'N/A',
          phone: item.phone || '9876543210',
          email: item.email || 'N/A',
          address: item.address || 'Essur - 603301',
          previousSchool: item.previousSchool || 'None',
          submittedAt: item.submittedAt || new Date().toISOString(),
          status: item.status || 'Pending Review',
        };
        const existingIndex = admissionsDb.findIndex((entry) => entry.id === id);
        if (existingIndex >= 0) {
          admissionsDb[existingIndex] = admission;
        } else {
          admissionsDb.unshift(admission);
        }
      }
      persistDatabase();
      res.json({ success: true, importedCount: records.length, collection: 'admissions' });
      return;
    }

    if (collection === 'fees' || collection === 'receipts') {
      for (const item of records) {
        const receiptNo = String(item.receiptNo || item.id || `WNS-RCP-${Date.now()}-${Math.random().toString().slice(2, 6)}`);
        const receipt: FeeReceipt = {
          id: item.id || receiptNo,
          receiptNo,
          studentName: item.studentName || 'Student',
          studentId: item.studentId || 'WNS-GUEST',
          grade: item.grade || 'General',
          term: item.term || 'Tuition Fee',
          amount: Number(item.amount || 0),
          paymentMethod: item.paymentMethod || 'UPI Transfer',
          transactionRef: item.transactionRef || `UPI-${Date.now().toString().slice(-8)}`,
          date: item.date || new Date().toLocaleDateString('en-IN'),
          status: item.status || 'VERIFIED',
        };
        const existingIndex = receiptsDb.findIndex((entry) => (entry.receiptNo || entry.id) === receiptNo || (entry.id && entry.id === receipt.id));
        if (existingIndex >= 0) {
          receiptsDb[existingIndex] = receipt;
        } else {
          receiptsDb.unshift(receipt);
        }
      }
      persistDatabase();
      res.json({ success: true, importedCount: records.length, collection: 'fees' });
      return;
    }

    res.status(400).json({ success: false, message: 'Unsupported import collection. Use students, admissions, or fees.' });
  } catch (error) {
    console.error('Import error:', error);
    res.status(500).json({ success: false, message: 'Import failed. Please check the input format.' });
  }
});

// Admin Batch Upload Students from Excel/CSV API
app.post('/api/admin/students/batch-upload', requireAdminAuth, (req, res) => {
  const { students } = req.body || {};

  if (!Array.isArray(students) || students.length === 0) {
    res.status(400).json({ success: false, message: 'No student records provided in payload.' });
    return;
  }

  let importedCount = 0;
  students.forEach((st: any, idx: number) => {
    const studentId = st.id || `WNS-2026-${String(10 + Object.keys(studentProfiles).length + idx).padStart(2, '0')}`;
    studentProfiles[studentId] = {
      id: studentId,
      name: st.name || 'Student',
      grade: st.grade || 'Class 3',
      rollNo: st.rollNo || String(idx + 1),
      phone: st.phone || '9876543210',
      dob: st.dob || '2018-06-15',
      fatherName: st.fatherName || 'Parent',
      motherName: st.motherName || 'Parent',
      classTeacher: st.classTeacher || 'Mrs. Anita Ramesh',
      attendancePercentage: 95.0,
      totalDays: 85,
      presentDays: 81,
      bloodGroup: st.bloodGroup || 'O+',
      feeStatus: st.feeStatus || 'Paid for Term 1',
      importantNotes: st.importantTag || 'Excel Import Record',
      term1Marks: [
        { subject: 'Tamil', mark: 90, max: 100, grade: 'A+' },
        { subject: 'English', mark: 88, max: 100, grade: 'A' },
        { subject: 'Mathematics', mark: 92, max: 100, grade: 'A+' },
        { subject: 'Science', mark: 91, max: 100, grade: 'A+' },
      ],
      homework: [
        { id: `hw-imp-${idx}`, subject: 'General', title: 'Welcome to Wisdom Student Portal!', dueDate: 'Next Week', status: 'Pending' },
      ],
      teacherNotes: `Record imported via Excel Batch Upload on ${new Date().toLocaleDateString('en-IN')}`,
    };
    importedCount++;
  });

  persistDatabase();

  res.json({
    success: true,
    importedCount,
    totalStudents: Object.keys(studentProfiles).length,
    message: `Successfully batch imported ${importedCount} student records from Excel file!`,
    students: Object.values(studentProfiles),
  });
});

// In-memory fee reminders log
interface FeeReminderLog {
  id: string;
  studentId: string;
  studentName: string;
  parentPhone: string;
  term: string;
  amount: number;
  channel: 'whatsapp' | 'sms' | 'email' | 'push';
  sentAt: string;
  status: 'DELIVERED' | 'SENT';
  messageText: string;
}

// Send Automated Fee Payment Reminder API
app.post('/api/student/fee-reminders/send', (req, res) => {
  const { studentId, studentName, parentPhone, channel, term, amount } = req.body || {};

  const cleanPhone = String(parentPhone || '9876543210').replace(/\D/g, '');
  const validChannels = ['whatsapp', 'sms', 'email', 'push'];
  const chosenChannel: 'whatsapp' | 'sms' | 'email' | 'push' = validChannels.includes(channel) ? channel : 'push';
  const feeTerm = term || 'Term 2 Tuition Fee';
  const feeAmount = Number(amount) || 3500;
  const name = studentName || 'Student';

  let messageText = '';
  if (chosenChannel === 'push') {
    messageText = `🔔 [Wisdom School Push Alert] Gentle Fee Reminder: ${feeTerm} ₹${feeAmount.toLocaleString('en-IN')} for ${name} is due in 3 days (15-Aug-2026). Click to pay instantly via UPI.`;
  } else if (chosenChannel === 'whatsapp') {
    messageText = `🔔 [Wisdom Nursery & Primary School, Essur]\nAutomated Fee Payment Reminder for ${name} (${feeTerm}).\n\nFee Amount: ₹${feeAmount.toLocaleString('en-IN')}\nDue Date: 15-Aug-2026\nSchool UPI ID: rsaravanan102002-1@okhdfcbank\n\nClick to Pay Instant UPI: https://wisdomessur.edu.in/fees?studentId=${studentId || 'WNS-2026-01'}\n\nHelpdesk: +91 9176593129`;
  } else {
    messageText = `WNS ESSUR: Fee Reminder for ${name}. ${feeTerm} ₹${feeAmount} due in 3 days (15-Aug-2026). Pay online: https://wisdomessur.edu.in/fees or school office. Ph: 9176593129`;
  }

  const newLog: FeeReminderLog = {
    id: `REM-${Date.now().toString().slice(-6)}`,
    studentId: studentId || 'WNS-2026-01',
    studentName: name,
    parentPhone: cleanPhone,
    term: feeTerm,
    amount: feeAmount,
    channel: chosenChannel,
    sentAt: new Date().toLocaleString('en-IN'),
    status: 'DELIVERED',
    messageText,
  };

  feeRemindersDb.unshift(newLog);
  persistDatabase();

  res.json({
    success: true,
    message: `Automated ${chosenChannel.toUpperCase()} Fee Reminder dispatched successfully`,
    reminder: newLog,
  });
});

// Bulk Send Fee Payment Reminders API (Multi-channel: WhatsApp, SMS, Email, Push)
app.post('/api/student/fee-reminders/bulk-send', (req, res) => {
  const { targets, channels, customMessage } = req.body || {};

  if (!Array.isArray(targets) || targets.length === 0) {
    res.status(400).json({ success: false, message: 'No target students selected for bulk sending.' });
    return;
  }

  const selectedChannels: ('whatsapp' | 'sms' | 'email' | 'push')[] = Array.isArray(channels) && channels.length > 0 
    ? channels 
    : ['whatsapp', 'sms', 'push'];

  const dispatchedLogs: FeeReminderLog[] = [];
  let whatsappCount = 0;
  let smsCount = 0;
  let emailCount = 0;
  let pushCount = 0;

  targets.forEach((target: any) => {
    const name = target.studentName || target.name || 'Student';
    const studentId = target.studentId || target.id || 'WNS-STUDENT';
    const phone = String(target.parentPhone || target.phone || '9876543210').replace(/\D/g, '');
    const term = target.term || target.feeParticulars || 'Term Fee';
    const amount = Number(target.amount || target.amountDue || 3500);
    const dueDate = target.dueDate || '15-Aug-2026';

    selectedChannels.forEach((ch) => {
      let msg = '';
      if (customMessage) {
        msg = customMessage
          .replace(/\{student_name\}/g, name)
          .replace(/\{student_id\}/g, studentId)
          .replace(/\{fee_term\}/g, term)
          .replace(/\{amount\}/g, `₹${amount.toLocaleString('en-IN')}`)
          .replace(/\{due_date\}/g, dueDate);
      } else if (ch === 'whatsapp') {
        msg = `🔔 [Wisdom Nursery & Primary School, Essur]\nOfficial Fee Due Reminder for ${name} (${studentId}, ${target.grade || 'Class'}).\n\nParticulars: ${term}\nAmount Payable: ₹${amount.toLocaleString('en-IN')}\nDue Date: ${dueDate}\nSchool UPI ID: rsaravanan102002-1@okhdfcbank\n\nPay online instantly: https://wisdomessur.edu.in/fees?studentId=${studentId}\nHelpdesk Ph: +91 9176593129`;
        whatsappCount++;
      } else if (ch === 'sms') {
        msg = `WISDOM SCHOOL ESSUR: Fee due notice for ${name}. ${term} ₹${amount} due on ${dueDate}. Pay online: https://wisdomessur.edu.in/fees or school office. Ph: 9176593129`;
        smsCount++;
      } else if (ch === 'email') {
        msg = `Subject: Wisdom School Fee Payment Notice - ${name} (${studentId})\nDear Parent, gentle reminder that ${term} of ₹${amount} for ${name} is due on ${dueDate}. Please settle via UPI or school cash counter.`;
        emailCount++;
      } else {
        msg = `🔔 [Wisdom Push Alert] Fee Reminder for ${name}: ${term} ₹${amount} due on ${dueDate}. Tap to pay.`;
        pushCount++;
      }

      const log: FeeReminderLog = {
        id: `REM-${Date.now().toString().slice(-6)}-${Math.floor(Math.random() * 1000)}`,
        studentId,
        studentName: name,
        parentPhone: phone,
        term,
        amount,
        channel: ch,
        sentAt: new Date().toLocaleString('en-IN'),
        status: 'DELIVERED',
        messageText: msg,
      };

      dispatchedLogs.push(log);
      feeRemindersDb.unshift(log);
    });
  });

  persistDatabase();

  res.json({
    success: true,
    message: `Bulk reminder campaign dispatched across ${selectedChannels.join(', ').toUpperCase()} to ${targets.length} parents!`,
    totalSent: dispatchedLogs.length,
    channelBreakdown: {
      whatsapp: whatsappCount,
      sms: smsCount,
      email: emailCount,
      push: pushCount,
    },
    logs: dispatchedLogs,
  });
});

// Campaign History Storage for Admin Notification Center

// GET Overdue Fee List API for Admin Notification Center
app.get('/api/admin/notifications/overdue-list', requireAdminAuth, (req, res) => {
  const overdueList = [
    {
      studentId: 'WNS-2026-03',
      studentName: 'M. Yogesh',
      grade: 'Class 5 - A',
      parentName: 'M. Manikandan',
      parentPhone: '9876543212',
      feeParticulars: 'Term 2 Tuition Fee',
      amountDue: 4200,
      dueDate: '15-Jul-2026',
      daysOverdue: 26,
      lastReminderSent: '05-Aug-2026',
      status: 'OVERDUE',
    },
    {
      studentId: 'WNS-2026-05',
      studentName: 'A. Dhruv',
      grade: 'Class 3 - A',
      parentName: 'A. Anand',
      parentPhone: '9876543215',
      feeParticulars: 'Term 2 Tuition Fee',
      amountDue: 3800,
      dueDate: '20-Jul-2026',
      daysOverdue: 21,
      lastReminderSent: '02-Aug-2026',
      status: 'OVERDUE',
    },
    {
      studentId: 'WNS-2026-06',
      studentName: 'M. Kavya',
      grade: 'Class 4 - B',
      parentName: 'M. Murali',
      parentPhone: '9876543216',
      feeParticulars: 'Term 2 & Transport Route 2',
      amountDue: 5200,
      dueDate: '01-Jul-2026',
      daysOverdue: 40,
      lastReminderSent: '28-Jul-2026',
      status: 'CRITICAL_OVERDUE',
    },
    {
      studentId: 'WNS-2026-07',
      studentName: 'V. Rahul',
      grade: 'Class 2 - B',
      parentName: 'V. Venkatesh',
      parentPhone: '9876543217',
      feeParticulars: 'Quarterly Installment #2',
      amountDue: 3200,
      dueDate: '10-Jul-2026',
      daysOverdue: 31,
      lastReminderSent: '01-Aug-2026',
      status: 'OVERDUE',
    },
    {
      studentId: 'WNS-2026-08',
      studentName: 'S. Bhuvanesh',
      grade: 'Class 1 - A',
      parentName: 'S. Selvam',
      parentPhone: '9876543218',
      feeParticulars: 'Term 2 & Book Set Dues',
      amountDue: 4500,
      dueDate: '05-Jul-2026',
      daysOverdue: 36,
      lastReminderSent: '30-Jul-2026',
      status: 'CRITICAL_OVERDUE',
    },
    {
      studentId: 'WNS-2026-04',
      studentName: 'K. Ananya',
      grade: 'LKG - Blossom',
      parentName: 'Karthik Raja',
      parentPhone: '9876543210',
      feeParticulars: 'Uniform & Accessories Fee',
      amountDue: 2100,
      dueDate: '25-Jul-2026',
      daysOverdue: 16,
      lastReminderSent: '04-Aug-2026',
      status: 'OVERDUE',
    },
  ];

  res.json({
    success: true,
    totalOverdueCount: overdueList.length,
    totalOverdueAmount: overdueList.reduce((sum, item) => sum + item.amountDue, 0),
    students: overdueList,
    campaignHistory: notificationCampaignHistoryDb,
  });
});

// Trigger One-Click Bulk Reminder Campaign API
app.post('/api/admin/notifications/bulk-reminder', requireAdminAuth, (req, res) => {
  const { channel, selectedStudentIds, customMessage, campaignTitle } = req.body || {};

  const targetIds: string[] = Array.isArray(selectedStudentIds) ? selectedStudentIds : [];
  if (targetIds.length === 0) {
    res.status(400).json({ success: false, message: 'Please select at least one parent/student to send reminders.' });
    return;
  }

  const newCampaign = {
    id: `CMP-${Date.now().toString().slice(-6)}`,
    title: campaignTitle || `Bulk Fee Reminder (${channel || 'WhatsApp + SMS'})`,
    channel: channel === 'whatsapp' ? 'WhatsApp' : channel === 'sms' ? 'SMS' : 'WhatsApp + SMS',
    recipientCount: targetIds.length,
    totalAmountOverdue: targetIds.length * 3800, // estimated
    messagePreview: customMessage || 'Wisdom School Fee Reminder Notice to parents',
    sentAt: new Date().toLocaleString('en-IN'),
    sentBy: 'Admin R. Saravanan',
    status: 'Delivered (100%)',
  };

  notificationCampaignHistoryDb.unshift(newCampaign);
  persistDatabase();

  res.json({
    success: true,
    message: `One-Click Bulk Reminders successfully sent via ${newCampaign.channel} to ${targetIds.length} parents!`,
    campaign: newCampaign,
    logs: notificationCampaignHistoryDb,
  });
});

// Fetch Fee Categories API
app.get('/api/fees/categories', (req, res) => {
  res.json({
    success: true,
    categories: feeCategoriesDb,
  });
});

// Update or Create Fee Category API
app.post('/api/fees/categories', (req, res) => {
  const { id, name, amount, description, categoryType, frequency } = req.body || {};

  if (!name || amount === undefined) {
    res.status(400).json({ success: false, message: 'Fee Name and Amount are required.' });
    return;
  }

  const existingIndex = feeCategoriesDb.findIndex((c) => c.id === id);
  if (existingIndex !== -1) {
    feeCategoriesDb[existingIndex] = {
      ...feeCategoriesDb[existingIndex],
      name: String(name).trim(),
      amount: Number(amount),
      description: String(description || '').trim(),
      categoryType: categoryType || feeCategoriesDb[existingIndex].categoryType,
      frequency: frequency || feeCategoriesDb[existingIndex].frequency,
    };
    persistDatabase();
    res.json({
      success: true,
      message: `Successfully updated fee category "${name}"`,
      category: feeCategoriesDb[existingIndex],
      categories: feeCategoriesDb,
    });
  } else {
    const newId = id || `fee-custom-${Date.now()}`;
    const newCat = {
      id: newId,
      key: newId.replace(/[^a-zA-Z0-9]/g, '_').toLowerCase(),
      name: String(name).trim(),
      amount: Number(amount),
      description: String(description || '').trim(),
      deletable: true,
      editable: true,
      categoryType: categoryType || 'Other',
      frequency: frequency || 'One-Time',
    };
    feeCategoriesDb.push(newCat);
    persistDatabase();
    res.json({
      success: true,
      message: `Successfully created new fee category "${name}"`,
      category: newCat,
      categories: feeCategoriesDb,
    });
  }
});

// Delete Fee Category API
app.delete('/api/fees/categories/:id', (req, res) => {
  const { id } = req.params;
  const target = feeCategoriesDb.find((c) => c.id === id);

  if (!target) {
    res.status(404).json({ success: false, message: 'Fee category not found.' });
    return;
  }

  feeCategoriesDb = feeCategoriesDb.filter((c) => c.id !== id);
  persistDatabase();
  res.json({
    success: true,
    message: `Successfully deleted fee category "${target.name}"`,
    categories: feeCategoriesDb,
  });
});

// Fetch Fee Reminders History API
app.get('/api/student/fee-reminders/:studentId', (req, res) => {
  const { studentId } = req.params;
  const history = feeRemindersDb.filter((r) => !studentId || r.studentId === studentId || studentId === 'all');
  res.json({
    success: true,
    history,
    upcomingSchedules: [
      {
        term: 'Term 2 Tuition & Activity Fee',
        grade: 'Class 3-A',
        dueDate: '15 August 2026',
        amount: 3500,
        status: 'UPCOMING_DUE',
        autoReminderSchedule: ['7 Days Prior (08-Aug-2026)', '3 Days Prior (12-Aug-2026)', 'On Due Date (15-Aug-2026)'],
      },
      {
        term: 'School Van Transport Fee (Term 2)',
        grade: 'Route 2 - Essur/Cheyyur',
        dueDate: '10 August 2026',
        amount: 1200,
        status: 'UPCOMING_DUE',
        autoReminderSchedule: ['5 Days Prior (05-Aug-2026)', 'On Due Date (10-Aug-2026)'],
      },
    ],
  });
});

// Update Admission Status API (Admin)
app.post('/api/admin/admission/status', requireAdminAuth, (req, res) => {
  const { id, status } = req.body;
  const target = admissionsDb.find((a) => a.id === id);
  if (target) {
    target.status = status;
    persistDatabase();
    res.json({ success: true, message: `Admission ${id} status updated to "${status}"`, application: target });
  } else {
    res.status(404).json({ success: false, message: 'Admission application not found' });
  }
});

// Admissions Submission API
app.post('/api/admissions', (req, res) => {
  const { studentName, dob, gender, grade, fatherName, motherName, phone, email, address, previousSchool } = req.body;

  if (!studentName || !grade || !phone || !fatherName) {
    res.status(400).json({ success: false, message: 'Please complete all required fields (Student Name, Grade, Father Name, Mobile Number).' });
    return;
  }

  const newApp: AdmissionApplication = {
    id: `ADM-ESSUR-2026-${Math.floor(100 + Math.random() * 900)}`,
    studentName,
    dob: dob || 'N/A',
    gender: gender || 'Unspecified',
    grade,
    fatherName,
    motherName: motherName || 'N/A',
    phone,
    email: email || 'N/A',
    address: address || 'Essur - 603301',
    previousSchool: previousSchool || 'None',
    submittedAt: new Date().toISOString(),
    status: 'Pending Review',
  };

  admissionsDb.unshift(newApp);
  persistDatabase();

  res.json({
    success: true,
    message: 'Admission Application Submitted Successfully!',
    application: newApp,
  });
});

// Get List of Admissions (for demo tracking)
app.get('/api/admissions', (req, res) => {
  res.json({ success: true, count: admissionsDb.length, applications: admissionsDb });
});

// Student Records CRUD APIs
app.get('/api/students', (req, res) => {
  res.json({
    success: true,
    count: Object.keys(studentProfiles).length,
    students: Object.values(studentProfiles),
  });
});

app.get('/api/students/:id', (req, res) => {
  const { id } = req.params;
  const student = studentProfiles[id];
  if (!student) {
    res.status(404).json({ success: false, message: 'Student record not found.' });
    return;
  }
  res.json({ success: true, student });
});

app.post('/api/students', requireAdminAuth, (req, res) => {
  const payload = req.body || {};
  const id = String(payload.studentId || payload.id || `WNS-${Date.now()}`);
  const student = {
    id,
    name: payload.name || 'Student',
    grade: payload.grade || 'Class 1',
    rollNo: payload.rollNo || payload.roll_no || '0',
    phone: payload.phone || '9876543210',
    dob: payload.dob || '2018-06-15',
    fatherName: payload.fatherName || 'Parent',
    motherName: payload.motherName || 'Parent',
    classTeacher: payload.classTeacher || 'Class Teacher',
    attendancePercentage: Number(payload.attendancePercentage || 0),
    totalDays: Number(payload.totalDays || 0),
    presentDays: Number(payload.presentDays || 0),
    bloodGroup: payload.bloodGroup || 'O+',
    feeStatus: payload.feeStatus || 'Pending',
    teacherNotes: payload.teacherNotes || '',
    importantNotes: payload.importantNotes || '',
    term1Marks: Array.isArray(payload.term1Marks) ? payload.term1Marks : [],
    homework: Array.isArray(payload.homework) ? payload.homework : [],
  };

  studentProfiles[id] = student;
  persistDatabase();
  res.json({ success: true, message: 'Student record saved.', student });
});

app.put('/api/students/:id', requireAdminAuth, (req, res) => {
  const { id } = req.params;
  const payload = req.body || {};
  const current = studentProfiles[id];
  if (!current) {
    res.status(404).json({ success: false, message: 'Student record not found.' });
    return;
  }

  const updated = {
    ...current,
    ...payload,
    id,
    name: payload.name || current.name,
    grade: payload.grade || current.grade,
    rollNo: payload.rollNo || payload.roll_no || current.rollNo,
    phone: payload.phone || current.phone,
    dob: payload.dob || current.dob,
    fatherName: payload.fatherName || current.fatherName,
    motherName: payload.motherName || current.motherName,
    classTeacher: payload.classTeacher || current.classTeacher,
    attendancePercentage: Number(payload.attendancePercentage ?? current.attendancePercentage ?? 0),
    totalDays: Number(payload.totalDays ?? current.totalDays ?? 0),
    presentDays: Number(payload.presentDays ?? current.presentDays ?? 0),
    bloodGroup: payload.bloodGroup || current.bloodGroup,
    feeStatus: payload.feeStatus || current.feeStatus,
    teacherNotes: payload.teacherNotes ?? current.teacherNotes,
    importantNotes: payload.importantNotes ?? current.importantNotes,
    term1Marks: Array.isArray(payload.term1Marks) ? payload.term1Marks : current.term1Marks || [],
    homework: Array.isArray(payload.homework) ? payload.homework : current.homework || [],
  };

  studentProfiles[id] = updated;
  persistDatabase();
  res.json({ success: true, message: 'Student record updated.', student: updated });
});

app.delete('/api/students/:id', requireAdminAuth, (req, res) => {
  const { id } = req.params;
  if (!studentProfiles[id]) {
    res.status(404).json({ success: false, message: 'Student record not found.' });
    return;
  }

  delete studentProfiles[id];
  const db = getDb();
  db.prepare('DELETE FROM student_records WHERE student_id = ?').run(id);
  persistDatabase();
  res.json({ success: true, message: `Student ${id} deleted.` });
});

// Submit Online Fee Payment & Generate Receipt API
app.post('/api/fees/pay', (req, res) => {
  const { studentName, studentId, grade, term, amount, paymentMethod, transactionRef } = req.body;

  if (!studentName || !amount) {
    res.status(400).json({ success: false, message: 'Student Name and Fee Amount are required.' });
    return;
  }

  const newReceipt: FeeReceipt = {
    receiptNo: `WNS-RCP-2026-${Math.floor(1000 + Math.random() * 9000)}`,
    studentName,
    studentId: studentId || 'WNS-GUEST',
    grade: grade || 'General',
    term: term || 'Tuition Fee',
    amount: Number(amount),
    paymentMethod: paymentMethod || 'UPI Transfer',
    transactionRef: transactionRef || `UPI-${Date.now().toString().slice(-8)}`,
    date: new Date().toLocaleDateString('en-IN'),
    status: 'VERIFIED',
  };

  receiptsDb.unshift(newReceipt);
  persistDatabase();

  res.json({
    success: true,
    message: 'Fee Payment Registered & Official Receipt Generated!',
    receipt: newReceipt,
  });
});

// Get Receipts
app.get('/api/fees/receipts', (req, res) => {
  res.json({ success: true, receipts: receiptsDb });
});

// AI Chatbot endpoint for parents/students
app.post('/api/ai/chat', async (req, res) => {
  try {
    const { message, history } = req.body;
    if (!message) {
      res.status(400).json({ success: false, error: 'Message text is required' });
      return;
    }

    const aiClient = getGeminiClient();
    if (!aiClient) {
      // Fallback friendly reply if API key is not configured yet
      res.json({
        success: true,
        reply: `Welcome to Wisdom Nursery and Primary School, Essur! 🏫\n\nAdministrator: R. Saravanan\nPhone/WhatsApp: +91 9176593129\nEmail: wisdomrs.tamil@gmail.com\n\nFor admissions, fees payment via GPay (UPI: rsaravanan102002-1@okhdfcbank), or school visits, please call us directly or submit your inquiry online!`,
      });
      return;
    }

    const systemInstruction = `You are "Wisdom AI Assistant", the official AI virtual counselor for Wisdom Nursery and Primary School located in Essur - 603301, Tamil Nadu, India.
Key Information:
- School Name: Wisdom Nursery and Primary School
- Location: Essur - 603301, Kanchipuram / Chengalpattu District, Tamil Nadu.
- Motto: "LEARN TODAY, LEAD TOMORROW."
- School Administrator: R. SARAVANAN
- Contact Mobile / WhatsApp: +91 9176593129
- Contact Email: wisdomrs.tamil@gmail.com
- Official UPI ID for Fee Payment: rsaravanan102002-1@okhdfcbank (Account Holder: R Saravanan)
- Classes Offered: Pre-Nursery, LKG, UKG, Standard 1, Standard 2, Standard 3, Standard 4, Standard 5.
- Curriculum: Tamil Nadu State Samacheer Kalvi (State Board) with strong English medium immersion, Tamil language foundation, Vedic Math, Science experiments, Computer skills, and Value education.
- School Facilities: Well-lit airy classrooms, digital smart learning aids, playground & sports equipment, clean drinking water, activity room, computer lab, safe transport guidance, and regular health checkups.
- Timings: Monday to Friday (8:30 AM to 3:30 PM), Saturday (8:30 AM to 12:30 PM for Primary).
- Admission Process: Simple online registration or school office visit, birth certificate copy, passport size photos, and Aadhaar card copy.
- Response Style: Extremely warm, respectful, polite, encouraging, and informative. Answer in clear English (or Tamil if the user queries in Tamil).`;

    const contents: any[] = [];
    if (history && Array.isArray(history)) {
      for (const msg of history) {
        contents.push({
          role: msg.role === 'assistant' ? 'model' : 'user',
          parts: [{ text: msg.text }],
        });
      }
    }
    contents.push({
      role: 'user',
      parts: [{ text: message }],
    });

    const response = await aiClient.models.generateContent({
      model: 'gemini-3.6-flash',
      contents: contents,
      config: {
        systemInstruction: systemInstruction,
        temperature: 0.7,
      },
    });

    const replyText = response.text || 'Thank you for contacting Wisdom Nursery and Primary School, Essur! How else can I assist you?';

    res.json({
      success: true,
      reply: replyText,
    });
  } catch (err: any) {
    console.error('Gemini API Error:', err);
    res.status(500).json({
      success: false,
      error: 'Unable to connect to AI Assistant right now.',
      fallbackReply: 'Wisdom Nursery and Primary School, Essur - 603301. Contact Admin R. Saravanan at +91 9176593129 for quick assistance.',
    });
  }
});

async function startServer() {
  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Wisdom School Server listening on http://localhost:${PORT}`);
  });
}

startServer();
