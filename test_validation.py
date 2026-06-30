#!/usr/bin/env python3
"""Test validation and output generation for both .xlsx and .xls formats."""
import io
import openpyxl
import xlwt
from app import app

client = app.test_client()

# Create a test .xlsx file
xlsx_buffer = io.BytesIO()
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Amount', 'BeneficiaryBankCode', 'BeneficiaryIBAN', 'BeneficiaryName', 'BeneficiaryCode', 'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryMobile', 'BeneficiaryIdentificationType', 'BeneficiaryIdentificationNo', 'ProductTypeCode'])
ws.append(['100', '001', 'PK36SCBL0000001123456702', 'Ali Khan', 'B1', 'R1', 'R2', 'ali@example.com', '03112753114', 'CNIC', '4220140603751', 'P1'])
ws.append(['200', '001', 'INVALID', 'Bad IBAN', 'B2', 'R1', 'R2', 'bad@example.com', '03112753115', 'CNIC', '4220140603752', 'P1'])  # Bad IBAN (too short and not uppercase)
wb.save(xlsx_buffer)
xlsx_bytes = xlsx_buffer.getvalue()

# Create a test .xls file  
xls_buffer = io.BytesIO()
book = xlwt.Workbook()
sheet = book.add_sheet('Sheet1')
headers = ['Amount', 'BeneficiaryBankCode', 'BeneficiaryIBAN', 'BeneficiaryName', 'BeneficiaryCode', 'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryMobile', 'BeneficiaryIdentificationType', 'BeneficiaryIdentificationNo', 'ProductTypeCode']
for col_idx, header in enumerate(headers):
    sheet.write(0, col_idx, header)
sheet.write(1, 0, 100)
sheet.write(1, 1, '001')
sheet.write(1, 2, 'PK36SCBL0000001123456702')
sheet.write(1, 3, 'Ali Khan')
sheet.write(1, 4, 'B1')
sheet.write(1, 5, 'R1')
sheet.write(1, 6, 'R2')
sheet.write(1, 7, 'ali@example.com')
sheet.write(1, 8, '03112753114')
sheet.write(1, 9, 'CNIC')
sheet.write(1, 10, '4220140603751')
sheet.write(1, 11, 'P1')
sheet.write(2, 0, 200)
sheet.write(2, 1, '001')
sheet.write(2, 2, 'INVALID')  # Bad IBAN (too short)
sheet.write(2, 3, 'Bad IBAN')
sheet.write(2, 4, 'B2')
sheet.write(2, 5, 'R1')
sheet.write(2, 6, 'R2')
sheet.write(2, 7, 'bad@example.com')
sheet.write(2, 8, '03112753115')
sheet.write(2, 9, 'CNIC')
sheet.write(2, 10, '4220140603752')
sheet.write(2, 11, 'P1')
book.save(xls_buffer)
xls_bytes = xls_buffer.getvalue()

print('[VALIDATION_TEST] Testing .xlsx validation')
response = client.post('/validate', data={'file': (io.BytesIO(xlsx_bytes), 'test.xlsx'), 'profile': 'raast'})
print(f'Status: {response.status_code}')
data = response.get_json()
print(f'Total records: {data["summary"]["total"]}')
print(f'Passed: {data["summary"]["passed"]}')
print(f'Rejected: {data["summary"]["rejected"]}')
print(f'File ID: {data.get("file_id")}')
assert response.status_code == 200, f"Expected 200, got {response.status_code}"
assert data['summary']['total'] == 2, "Expected 2 total records"
assert data['summary']['passed'] == 1, "Expected 1 passed record"
assert data['summary']['rejected'] == 1, "Expected 1 rejected record"

print('\n[VALIDATION_TEST] Testing .xls validation')
response = client.post('/validate', data={'file': (io.BytesIO(xls_bytes), 'test.xls'), 'profile': 'raast'})
print(f'Status: {response.status_code}')
data = response.get_json()
print(f'Total records: {data["summary"]["total"]}')
print(f'Passed: {data["summary"]["passed"]}')
print(f'Rejected: {data["summary"]["rejected"]}')
print(f'File ID: {data.get("file_id")}')
assert response.status_code == 200, f"Expected 200, got {response.status_code}"
assert data['summary']['total'] == 2, "Expected 2 total records"
assert data['summary']['passed'] == 1, "Expected 1 passed record"
assert data['summary']['rejected'] == 1, "Expected 1 rejected record"

print('\n[VALIDATION_TEST] Outputs generated correctly for both formats')
print('[VALIDATION_TEST] All validation tests passed!')
