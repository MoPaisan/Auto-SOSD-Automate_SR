# Auto-SOSD-Automate_SR

โปรเจกต์สำหรับอัตโนมัติการจัดการ SR Promotion Prepaid Main On top Data ปัญหา

## Table of Contents
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)

## Features
- Login ระบบ IM
- ดึง Report promotion ที่เว็บ IM
- แยกประเภท case reject
- เช็ค case marketing approve
- Login ระบบ MO
- อัพโหลดไฟล์ case marketing approve ที่ระบบ MO
- แยกข้อมูลเพื่อไปปิดเคส reject กับ complete
- อัพโหลดไฟล์ไฟล์เพื่อปิดเคสที่ ระบบ IM

## Installation
วิธีติดตั้ง/ตั้งค่าโปรเจกต์ เช่น
```sh
git clone https://github.com/yourname/project.git
cd project
npm install
```
หรือสำหรับ .NET
```sh
git clone https://github.com/yourname/project.git
cd project
dotnet restore
```

## Usage
วิธีรันหรือใช้งานโปรเจกต์ เช่น
```sh
dotnet run
```

## Configuration
- แก้ไขไฟล์ SOSD_AutomateSR.config ตามรายละเอียดที่ต้องการ
- ตัวอย่าง config
===============================================================================================
IM
===============================================================================================
URLIM => ลิงค์เว็บ IM ยกตัวอย่าง https://myim.intra.ais/prweb/PRWebLDAP1/
ResuIM => ชื่อผู้ใช้ ยกตัวอย่าง testsystem
SsapWordIM => รหัสผ่าน ยกตัวอย่าง !@#$1234
titleReport => หัวข้อรายงานที่ต้องการดาวน์โหลด ยกตัวอย่าง ReportSOSDPromotion_ServiceRequest
DayRange => จำนวนวันที่จะทำการดาวน์โหลดไฟล์ ยกตัวอย่าง 6
StatusCase => สถานะของเคสที่ต้องการดาวน์โหลด ยกตัวอย่าง Pending-Review
===============================================================================================
My Office
===============================================================================================
URLMyofficeWeb => ลิงค์เว็บ MO ยกตัวอย่าง https://myoffice-portal.intra.ais
ResuMyoffice => ชื่อผู้ใช้ ยกตัวอย่าง testsystem@mail.com
SsapWordMyoffice =>  รหัสผ่าน ยกตัวอย่าง !@#$1234
===============================================================================================
Path Folder
===============================================================================================
PathDownload => path ไฟล์ Download ยกตัวอย่าง C:\Users\profile\Downloads
PathReportSOSD => path ไฟล์ของ report ที่ download จาก IM ยกตัวอย่าง D:\SOSD_AutomateSR\ReportSOSDPromotion
PathOutputReport => path ไฟล์ที่เก็บไฟล์สำหรับปิดเคส ยกตัวอย่าง D:\SOSD_AutomateSR\ReportCloseCase
PathReportDaily => path ไฟล์ report สรุปรายวัน ยกตัวอย่าง D:\SOSD_AutomateSR\ReportSummary\ReportDaily
PathReportMonthly => path ไฟล์ report สรุปรายเดือน ยกตัวอย่าง D:\SOSD_AutomateSR\ReportSummary\ReportMonthly
PathTemplateCloseCase => path ไฟล์ Template ไฟล์สำหรับปิดเคส ยกตัวอย่าง D:\SOSD_AutomateSR\Templates\Import_CloseCaseIM_Template.xlsx
PathTemplateCheckMO => path ไฟล์ Template ไฟล์สำหรับ case marketing approve ยกตัวอย่าง D:\SOSD_AutomateSR\Templates\PSP_SR_Template.xlsx
PathMappingFile => path ไฟล์ Template ไฟล์สำหรับเช็ค FeatureCode ยกตัวอย่าง D:\SOSD_AutomateSR\Templates\Mapping_FeatureCode.xlsx
  


