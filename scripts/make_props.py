"""
Tool to regenerate tnefparse/properties.py

The mapping of property names to numeric value is maintained in data/properties.txt

Use:
    python scripts/make_props.py

"""
from collections import OrderedDict
from pathlib import Path

properties = OrderedDict()

for prop in Path('data/properties.txt').read_text().splitlines():
    prop = prop.strip()
    if not prop or prop.startswith('#'):
        continue
    name, val = prop.split(' ')
    properties[name] = val
sort_by_val = OrderedDict(sorted(properties.items(), key=lambda t: t[1]))

with Path('tnefparse/properties.py').open('w') as out:

    out.write('''"""
Property reference docs:

 - https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapping-canonical-property-names-to-mapi-names#tagged-properties
 - https://interoperability.blob.core.windows.net/files/MS-OXPROPS/[MS-OXPROPS].pdf
 - https://fossies.org/linux/libpst/xml/MAPI_definitions.pdf
 - https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/mapi-constants


+----------------+----------------+-------------------------------------------------------------------------------+
| Range minimum  | Range maximum  |                                 Description                                   |
+----------------+----------------+-------------------------------------------------------------------------------+
| 0x0001         | 0x0BFF         | Message object envelope property; reserved                                    |
| 0x0C00         | 0x0DFF         | Recipient property; reserved                                                  |
| 0x0E00         | 0x0FFF         | Non-transmittable Message property; reserved                                  |
| 0x1000         | 0x2FFF         | Message content property; reserved                                            |
| 0x3000         | 0x33FF         | Multi-purpose property that can appear on all or most objects; reserved       |
| 0x3400         | 0x35FF         | Message store property; reserved                                              |
| 0x3600         | 0x36FF         | Folder and address book container property; reserved                          |
| 0x3700         | 0x38FF         | Attachment property; reserved                                                 |
| 0x3900         | 0x39FF         | Address Book object property; reserved                                        |
| 0x3A00         | 0x3BFF         | Mail user object property; reserved                                           |
| 0x3C00         | 0x3CFF         | Distribution list property; reserved                                          |
| 0x3D00         | 0x3DFF         | Profile section property; reserved                                            |
| 0x3E00         | 0x3EFF         | Status object property; reserved                                              |
| 0x4000         | 0x57FF         | Transport-defined envelope property                                           |
| 0x5800         | 0x5FFF         | Transport-defined recipient property                                          |
| 0x6000         | 0x65FF         | User-defined non-transmittable property                                       |
| 0x6600         | 0x67FF         | Provider-defined internal non-transmittable property                          |
| 0x6800         | 0x7BFF         | Message class-defined content property                                        |
| 0x7C00         | 0x7FFF         | Message class-defined non-transmittable property                              |
| 0x8000         | 0xFFFF         | Reserved for mapping to named properties. The exceptions to this rule are     |
|                |                |   some of the address book tagged properties (those with names beginning with |
|                |                |   PIDTagAddressBook). Many are static property IDs but are in this range.     |
+----------------+----------------+-------------------------------------------------------------------------------+
"""


''')

    for name, val in sort_by_val.items():
        out.write(f'{name} = {val}\n')

    out.write('\n')
    out.write('CODE_TO_NAME = {\n')
    for name, _ in sort_by_val.items():
        out.write(f'    {name}: "{name}",\n')
    out.write('}\n')
