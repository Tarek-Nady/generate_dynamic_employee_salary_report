{
    'name': 'Employee Payroll Report',
    'version': '17.0.1.0.0',  # Update version to match your Odoo version
    'category': 'Human Resources',
    'summary': 'Basic Employee Salary Management',
    'description': """Manage basic employee salary information""",
    'author': 'Your Name',
    'depends': ['base', 'hr'],  # Add base explicitly
    'data': [
        'security/ir.model.access.csv',
        'views/employee_salary_views.xml',
    ],
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}