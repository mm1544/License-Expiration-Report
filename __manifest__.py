{
    'name': 'Licence Expiration Report',
    'version': '1.0',
    'category': 'Generic Modules/Others',
    'summary': 'Generates Licence/Software Expiration Report XLSX and sends to the designated person.',
    'sequence': '1',
    'author': 'Martynas Minskis',
    'depends': ['sale'],
    'demo': [],
    'data': [

        # Sequence: security, data, wizards, views
        'views/license_expiry_report.xml',
        'views/licence_length_months.xml',
    ],
    'demo': [],
    'qweb': [],

    'installable': True,
    'application': True,
    'auto_install': False,
    #     'licence': 'OPL-1',
}
