from flask import Blueprint, render_template, request, flash, redirect, url_for
import pandas as pd
from datetime import datetime
from app import db
from app.models import (
    Transaction, TeleSales, Renewal, RenewalData, ProductLookup, 
    Partner, PartnerContact, Contact, Company
)
import os
from werkzeug.utils import secure_filename
from flask import Flask, render_template

bp = Blueprint('main', __name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

@bp.route('/upload')
def upload():
    return render_template('upload.html')


def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@bp.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        
        file = request.files['file']
        file_type = request.form.get('file_type')
        
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, filename)
            
            if not os.path.exists(UPLOAD_FOLDER):
                os.makedirs(UPLOAD_FOLDER)
            
            file.save(filepath)
            
            try:
                process_excel(filepath, file_type)
                flash('File successfully uploaded and processed')
            except Exception as e:
                flash(f'Error processing file: {str(e)}')
                return redirect(request.url)
            
            return redirect(url_for('main.upload_file'))
    
    return render_template('upload.html')
def process_excel(filepath, file_type):
    if file_type == 'transaction':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            record = Transaction(
                currency=row.get('Currency'),
                location=row.get('Location'),
                region=row.get('Region'),
                sales_person=row.get('Sales Person'),
                customer_name=row.get('Customer Name'),
                product=row.get('Product'),
                nature_of_business=row.get('Nature of Business'),
                bu=row.get('BU'),
                partner_location=row.get('Partner Location'),
                partner=row.get('Partner'),
                type=row.get('Type'),
                psm=row.get('PSM'),
                partner_led=row.get('Partner Led?') == 'Yes',
                partner_account_manager_name=row.get('Partner Account Manager Name'),
                designation=row.get('Designation'),
                email_id=row.get('Email ID'),
                phone_number=row.get('Phone Number'),
                why_did_they_buy=row.get('Why did they buy?'),
                inv_date=pd.to_datetime(row.get('Inv date')),
                qtr=row.get('Qtr'),
                year=row.get('Year'),
                inv_value=row.get('Inv Value'),
                gp=row.get('GP'),
                comments=row.get('Comments')
            )
            db.session.add(record)
    
    elif file_type == 'tele_sales':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            # Handle NaN values by converting them to None or 0
            deals_closed = row.get('Deals Closed')
            if pd.isna(deals_closed):
                deals_closed = None  # or 0 if you prefer
            
            comments_notes = row.get('Comments/Notes')
            if pd.isna(comments_notes):
                comments_notes = None
            
            record = TeleSales(
                date=pd.to_datetime(row.get('Date')),
                rep_name=row.get('Rep Name'),
                total_calls_made=row.get('Total Calls Made', 0),
                new_calls=row.get('New Calls', 0),
                follow_up_calls=row.get('Follow Up Calls', 0),
                not_connected=row.get('Not connected', 0),
                connected_buy_not_interested=row.get('Connected buy Not Interested', 0),
                connected_and_asked_to_call_back=row.get('Connected and asked to call back', 0),
                connected_call=row.get('Connected Call', 0),
                emails_sent=row.get('Emails Sent', 0),
                new_emails=row.get('New Emails', 0),
                follow_up_emails=row.get('Follow Up Emails', 0),
                total_linkedin=row.get('Total LinkedIn', 0),
                linkedin_new_connect=row.get('LinkedIn New Connect', 0),
                linkedin_followups=row.get('LinkedIn Followups', 0),
                edms_sent=row.get("EDM's Sent", 0),
                appointments_set=row.get('Appointments Set', 0),
                demos_scheduled=row.get('Demos Scheduled', 0),
                meetings_held=row.get('Meetings Held', 0),
                deals_closed=deals_closed,
                notes_updated_in_crm=row.get('Notes Updated in CRM') == 'Yes',
                comments_notes=comments_notes
            )
            db.session.add(record)
    
    elif file_type == 'renewal':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            record = Renewal(
                date=pd.to_datetime(row.get('Date')),
                specialist_name=row.get('Specialist Name'),
                partners_touched=row.get('Partners Touched'),
                calls_made=row.get('Calls Made'),
                emails_sent=row.get('Emails Sent'),
                renewals_due=row.get('Renewals Due'),
                renewals_closed=row.get('Renewals Closed'),
                at_risk_accounts_engaged=row.get('At-Risk Accounts Engaged'),
                upsell_opportunities_identified=row.get('Upsell Opportunities Identified'),
                total_arr_renewed=float(str(row.get('Total ARR Renewed ($)')).replace('$', '').replace(',', '')) if pd.notna(row.get('Total ARR Renewed ($)')) else None,
                notes_updated_in_crm=row.get('Notes Updated in CRM') == 'Yes',
                churn_risk_notes=row.get('Churn Risk Notes'),
                productivity_comments=row.get('Productivity / Comments')
            )
            db.session.add(record)
    
    elif file_type == 'renewal_data':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            record = RenewalData(
                currency=row.get('Currency'),
                location=row.get('Location'),
                sales_person=row.get('Sales Person'),
                customer_name=row.get('Customer Name'),
                product=row.get('Product'),
                nature_of_business=row.get('Nature of Business'),
                bu=row.get('BU'),
                partner_location=row.get('Partner Location'),
                partner=row.get('Partner'),
                psm=row.get('PSM'),
                partner_account_manager_name=row.get('Partner Account Manager Name'),
                designation=row.get('Designation'),
                email_id=row.get('Email ID'),
                phone_number=row.get('Phone Number'),
                date_of_renewal=pd.to_datetime(row.get('Date of Renewal')) if pd.notna(row.get('Date of Renewal')) else None,
                last_year_invoice_date=pd.to_datetime(row.get('Last Year Invoice Date')) if pd.notna(row.get('Last Year Invoice Date')) else None,
                last_year_invoice_value=row.get('Last year Invoice Value'),
                last_year_margins=row.get('Last Year Margins'),
                this_year_technobind_price=row.get('This Year Technobind Price'),
                this_year_partner_price=row.get('This Year Partner Price'),
                status=row.get('Status'),
                comments=row.get('Comments')
            )
            db.session.add(record)
    
    elif file_type == 'product_lookup':
         df = pd.read_excel(filepath, sheet_name='Sheet1')
         for _, row in df.iterrows():
            record = ProductLookup(
                product_name=row.get('Product Name'),
                primary_industry_focus=row.get('Primary Industry Focus'),
                ideal_customer_profiles=row.get('Ideal Customer Profiles'),
                persona=row.get('Persona'),
                role=row.get('Role'),
                key_concerns=row.get('Key Concerns'),
                problem_statement=row.get('Problem Statement'),
                value_propositions=row.get('Value Propositions')
            )
            db.session.add(record)
    
    elif file_type == 'partner':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            record = Partner(
                company_name=row.get('Company Name'),
                partner_type=row.get('Partner Type'),
                website_url=row.get('Website URL'),
                headquarters_location=row.get('Headquarters Location'),
                hq_address=row.get('HQ Address'),
                regional_presence=row.get('Regional Presence'),
                partner_tier=row.get('Partner Tier'),
                top_oems=row.get("Top OEM's"),
                industry_focus=row.get('Industry Focus'),
                tech_stack_focus=row.get('Tech Stack Focus'),
                tech_stack_expertise=row.get('Tech Stack Expertise'),
                vendor_certifications=row.get('Vendor Certifications'),
                key_services_offered=row.get('Key Services Offered'),
                client_size_focus=row.get('Client Size Focus'),
                years_in_operation=row.get('Years in Operation'),
                number_of_employees=row.get('Number of Employees'),
                annual_revenue_est=row.get('Annual Revenue (Est.)'),
                contact_person_name=row.get('Contact Person Name'),
                contact_email=row.get('Contact Email'),
                contact_phone=row.get('Contact Phone'),
                linkedin_profile=row.get('LinkedIn Profile'),
                partner_status=row.get('Partner Status'),
                last_engagement_date=pd.to_datetime(row.get('Last Engagement Date')) if pd.notna(row.get('Last Engagement Date')) else None,
                notes_comments=row.get('Notes / Comments')
            )
            db.session.add(record)
    
    elif file_type == 'partner_contact':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            # Handle empty or NaN values
            decision_maker = row.get('Decision Maker?', '')
            dob = row.get('Date of Birth')
            
            record = PartnerContact(
                contact_name=row.get('Contact Name'),
                job_title=row.get('Job Title'),
                email_address=row.get('Email Address'),
                phone_number=row.get('Phone Number'),
                linkedin_profile=row.get('LinkedIn Profile'),
                company_name=row.get('Company Name'),
                company_website=row.get('Company Website'),
                company_tier=row.get('Company Tier'),
                department=row.get('Department'),
                location_city_country=row.get('Location (City/Country)'),
                primary_region=row.get('Primary Region'),
                products_handled=row.get('Products Handled'),
                is_decision_maker=decision_maker.upper() == 'Y' if pd.notna(decision_maker) else False,
                date_of_birth=pd.to_datetime(dob) if pd.notna(dob) else None,
                influence_level=row.get('Influence Level'),
                engagement_type=row.get('Engagement Type'),
                first_contact_date=pd.to_datetime(row.get('First Contact Date')) if pd.notna(row.get('First Contact Date')) else None,
                last_contact_date=pd.to_datetime(row.get('Last Contact Date')) if pd.notna(row.get('Last Contact Date')) else None,
                preferred_contact_method=row.get('Preferred Contact Method'),
                communication_status=row.get('Communication Status'),
                notes_history=row.get('Notes / History')
            )
            db.session.add(record)
    
    elif file_type == 'contact':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            record = Contact(
                organization_name=row.get('Organization Name'),
                organization_founded_year=int(row.get('Organization Founded Year')) if pd.notna(row.get('Organization Founded Year')) else None,
                organization_market_cap=row.get('Organization Market Cap'),
                phone_number_1=row.get('Phone Number 1'),
                phone_number_2=row.get('Phone Number 2'),
                phone_status=row.get('Phone Status'),
                organization_primary_domain=row.get('Organization Primary Domain'),
                city=row.get('City'),
                state=row.get('State'),
                country=row.get('Country'),
                person_name=row.get('Person Name'),
                first_name=row.get('First Name'),
                last_name=row.get('Last Name'),
                person_linkedin_url=row.get('Person Linkedin Url'),
                designation=row.get('Designation'),
                email_status=row.get('Email Status'),
                email=row.get('Email'),
                organization_facebook_url=row.get('Organization Facebook Url'),
                organization_linkedin_url=row.get('Organization Linkedin Url'),
                organization_twitter_url=row.get('Organization Twitter Url'),
                organization_website_url=row.get('Organization Website Url'),
                employee_size=row.get('Employee Size'),
                primary_industry=row.get('Primary Industry'),
                comments=row.get('Comments')
            )
            db.session.add(record)
    
    elif file_type == 'company':
        df = pd.read_excel(filepath, sheet_name='Sheet1')
        for _, row in df.iterrows():
            record = Company(
                company=row.get('Company'),
                head_office_location=row.get('Head Office Location'),
                primary_industry=row.get('Primary Industry'),
                industry=row.get('Industry'),
                sub_industry=row.get('Sub Industry'),
                type=row.get('Type'),
                location=row.get('Location'),
                employee_count=row.get('Employee Count'),
                revenue_range=row.get('Revenue Range'),
                num_employees=row.get('# Employees'),
                industry2=row.get('Industry2'),
                website=row.get('Website'),
                company_linkedin_url=row.get('Company Linkedin Url'),
                facebook_url=row.get('Facebook Url'),
                twitter_url=row.get('Twitter Url'),
                keywords=row.get('Keywords'),
                company_phone=row.get('Company Phone'),
                seo_description=row.get('SEO Description'),
                technologies=row.get('Technologies'),
                total_funding=row.get('Total Funding'),
                latest_funding=row.get('Latest Funding'),
                latest_funding_amount=row.get('Latest Funding Amount'),
                last_raised_at=pd.to_datetime(row.get('Last Raised At')) if pd.notna(row.get('Last Raised At')) else None,
                annual_revenue=row.get('Annual Revenue'),
                number_of_retail_locations=row.get('Number of Retail Locations'),
                short_description=row.get('Short Description'),
                founded_year=row.get('Founded Year'),
                comments=row.get('Comments')
            )
            db.session.add(record)
    
    db.session.commit()