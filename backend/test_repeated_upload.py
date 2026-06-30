#!/usr/bin/env python3
"""Test repeated upload flow without browser refresh."""
import io
import time
import openpyxl
import xlwt
from app import app

client = app.test_client()

def create_xlsx():
    """Create a valid test .xlsx file."""
    buffer = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(['Amount', 'BeneficiaryBankCode', 'BeneficiaryIBAN', 'BeneficiaryName', 'BeneficiaryCode', 'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryMobile', 'BeneficiaryIdentificationType', 'BeneficiaryIdentificationNo', 'ProductTypeCode'])
    ws.append(['100', '001', 'PK36SCBL0000001123456702', 'Ali Khan', 'B1', 'R1', 'R2', 'ali@example.com', '03112753114', 'CNIC', '4220140603751', 'P1'])
    wb.save(buffer)
    return buffer.getvalue()

def create_xls():
    """Create a valid test .xls file."""
    buffer = io.BytesIO()
    book = xlwt.Workbook()
    sheet = book.add_sheet('Sheet1')
    headers = ['Amount', 'BeneficiaryBankCode', 'BeneficiaryIBAN', 'BeneficiaryName', 'BeneficiaryCode', 'ReferenceField1', 'ReferenceField2', 'BeneficiaryEmail', 'BeneficiaryMobile', 'BeneficiaryIdentificationType', 'BeneficiaryIdentificationNo', 'ProductTypeCode']
    for col_idx, header in enumerate(headers):
        sheet.write(0, col_idx, header)
    sheet.write(1, 0, 200)
    sheet.write(1, 1, '001')
    sheet.write(1, 2, 'PK36SCBL0000001123456703')
    sheet.write(1, 3, 'Bob Smith')
    sheet.write(1, 4, 'B2')
    sheet.write(1, 5, 'R1')
    sheet.write(1, 6, 'R2')
    sheet.write(1, 7, 'bob@example.com')
    sheet.write(1, 8, '03112753115')
    sheet.write(1, 9, 'CNIC')
    sheet.write(1, 10, '4220140603752')
    sheet.write(1, 11, 'P1')
    book.save(buffer)
    return buffer.getvalue()

print('[REPEATED_UPLOAD_TEST] Testing unlimited file uploads in one session\n')

# Upload 1: .xlsx file
print('[TEST_1] Upload file A (.xlsx)')
response = client.post('/analyze', data={'file': (io.BytesIO(create_xlsx()), 'test_a.xlsx')})
assert response.status_code == 200, f'Expected 200, got {response.status_code}'
data1 = response.get_json()
assert data1.get('success') == True
print(f'✓ File A uploaded: {data1.get("row_count")} rows')

response = client.post('/validate', data={'file': (io.BytesIO(create_xlsx()), 'test_a.xlsx'), 'profile': 'raast'})
assert response.status_code == 200
result1 = response.get_json()
assert result1.get('file_id') is not None
print(f'✓ File A validated: {result1["summary"]["passed"]} passed, {result1["summary"]["rejected"]} rejected')

# Upload 2: Different .xls file
print('\n[TEST_2] Upload file B (.xls)')
response = client.post('/analyze', data={'file': (io.BytesIO(create_xls()), 'test_b.xls')})
assert response.status_code == 200
data2 = response.get_json()
assert data2.get('success') == True
print(f'✓ File B uploaded: {data2.get("row_count")} rows')

response = client.post('/validate', data={'file': (io.BytesIO(create_xls()), 'test_b.xls'), 'profile': 'raast'})
assert response.status_code == 200
result2 = response.get_json()
assert result2.get('file_id') is not None
print(f'✓ File B validated: {result2["summary"]["passed"]} passed, {result2["summary"]["rejected"]} rejected')

# Ensure file IDs are different
assert result1['file_id'] != result2['file_id'], 'File IDs should be different'
print(f'✓ File IDs are different (cache is working correctly)')

# Upload 3: Same file as first (should work)
print('\n[TEST_3] Upload same file A again')
response = client.post('/analyze', data={'file': (io.BytesIO(create_xlsx()), 'test_a.xlsx')})
assert response.status_code == 200
data3 = response.get_json()
assert data3.get('success') == True
print(f'✓ File A uploaded again: {data3.get("row_count")} rows')

response = client.post('/validate', data={'file': (io.BytesIO(create_xlsx()), 'test_a.xlsx'), 'profile': 'raast'})
assert response.status_code == 200
result3 = response.get_json()
assert result3.get('file_id') is not None
print(f'✓ File A validated again: {result3["summary"]["passed"]} passed, {result3["summary"]["rejected"]} rejected')

# Ensure this is a fresh file ID (new upload)
assert result3['file_id'] != result1['file_id'], 'New validation should have new file ID'
print(f'✓ New file ID created (not reusing previous validation)')

# Upload 4-10: Rapid repeated uploads
print('\n[TEST_4-10] Rapid repeated uploads (7 more times)')
for i in range(7):
    files = [create_xlsx(), create_xls()]
    filename = f'rapid_{i}.{"xlsx" if i % 2 == 0 else "xls"}'
    
    response = client.post('/analyze', data={'file': (io.BytesIO(files[i % 2]), filename)})
    assert response.status_code == 200
    
    response = client.post('/validate', data={'file': (io.BytesIO(files[i % 2]), filename), 'profile': 'raast'})
    assert response.status_code == 200
    result = response.get_json()
    assert result.get('file_id') is not None
    print(f'✓ Upload {i+1}: Validated successfully')

print('\n[REPEATED_UPLOAD_TEST] ✓ All 10 upload/validation cycles passed!')
print('[REPEATED_UPLOAD_TEST] No browser refresh required - state properly managed')
