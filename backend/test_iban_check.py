import json
from app import validate_rows

RULES = json.loads(open('rules.json', encoding='utf-8').read())

rows = [
    {
        'Amount': '100.50',
        'BeneficiaryBankCode': '001',
        'BeneficiaryIBAN': 'pk36scbl0000001123456709',
        'BeneficiaryName': 'IBAN Invalid',
        'BeneficiaryCode': 'B1',
        'ReferenceField1': 'R1',
        'ReferenceField2': 'R2',
        'BeneficiaryEmail': 'ali@example.com',
        'BeneficiaryMobile': '03112753114',
        'BeneficiaryIdentificationType': 'CNIC',
        'BeneficiaryIdentificationNo': '4220140603755',
        'ProductTypeCode': '',
    }
]

print(validate_rows(rows, 'raast', RULES, debug=True))
print(repr(rows[0]['BeneficiaryIBAN']))
print([ord(c) for c in rows[0]['BeneficiaryIBAN'][:4]])
