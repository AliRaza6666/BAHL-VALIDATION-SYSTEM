#!/usr/bin/env python3
"""End-to-end test for .xlsx and .xls compatibility."""
import io
import json
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
book.save(xls_buffer)
xls_bytes = xls_buffer.getvalue()

print('[E2E_TEST] Testing .xlsx upload')
response = client.post('/analyze', data={'file': (io.BytesIO(xlsx_bytes), 'test.xlsx')})
print(f'Status: {response.status_code}')
data = response.get_json()
print(f'Success: {data.get("success")}')
print(f'Rows: {data.get("row_count")}')
print(f'Profile: {data.get("profile")}')
assert response.status_code == 200, f"Expected 200, got {response.status_code}"
assert data.get('success') == True, "Expected success=true"
assert data.get('row_count') == 1, "Expected 1 row"
assert data.get('profile') == 'raast', "Expected raast profile detection"

print('\n[E2E_TEST] Testing .xls upload')
response = client.post('/analyze', data={'file': (io.BytesIO(xls_bytes), 'test.xls')})
print(f'Status: {response.status_code}')
data = response.get_json()
print(f'Success: {data.get("success")}')
print(f'Rows: {data.get("row_count")}')
print(f'Profile: {data.get("profile")}')
assert response.status_code == 200, f"Expected 200, got {response.status_code}"
assert data.get('success') == True, "Expected success=true"
assert data.get('row_count') == 1, "Expected 1 row"
assert data.get('profile') == 'raast', "Expected raast profile detection"

print('\n[E2E_TEST] Testing corrupt .xls')
response = client.post('/analyze', data={'file': (io.BytesIO(b'not-excel'), 'bad.xls')})
print(f'Status: {response.status_code}')
data = response.get_json()
print(f'Success: {data.get("success")}')
print(f'Message: {data.get("message")}')
assert response.status_code == 400, f"Expected 400, got {response.status_code}"
assert data.get('success') == False, "Expected success=false"
assert data.get('message') == 'Uploaded Excel file is corrupted.', f"Expected corrupted message, got {data.get('message')}"

print('\n[E2E_TEST] Testing .csv rejection')
response = client.post('/analyze', data={'file': (io.BytesIO(b'Amount,Name\n100,Ali'), 'test.csv')})
print(f'Status: {response.status_code}')
data = response.get_json()
print(f'Success: {data.get("success")}')
print(f'Message: {data.get("message")}')
assert response.status_code == 400, f"Expected 400, got {response.status_code}"
assert data.get('success') == False, "Expected success=false"

print('\n[E2E_TEST] All tests passed!')
