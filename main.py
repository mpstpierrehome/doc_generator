import datetime
# from datetime import date
import pandas as pd
from faker import Faker
# import pytz as pytz
from mailmerge import MailMerge
# from faker.providers import BaseProvider

faker = Faker(locale='en_US')
STATUS_LIST = ("Gold", "Bronze", "Silver")
template_1 = "templates/CustomerLetter.docx"
COMPANY_LIST = ("Stark Industries",
                "Wayne Enterprises",
                "Weyland-Yutani Corporation",
                "Umbrella Corporation",
                "Cyberdyne Systems",
                "InGen",
                "OCP",
                "Tyrell Corporation",
                "ACME",
                "Roxxon Energy Corporation",
                "Spectral Dynamics",
                "Blue Sun Corporation",
                "Palantine",
                "Umbrella Academy",
                "Brawndo",
                "LexCorp",
                "The Company",
                "The Corporation",
                "The Combine",
                "Aperture Science",
                )

PRODUCTS = ("Lasso", "Superman Cape", "Decks of Cards", "Carrots", "Salami")


def random_status() -> str:
    return faker.random_element(elements=STATUS_LIST)


def company_name() -> str:
    return faker.random_element(elements=COMPANY_LIST)


def product() -> str:
    return faker.random_element(elements=PRODUCTS)


def city() -> str:
    return faker.city()


def address() -> str:
    return faker.address()


def zipcode() -> str:
    return faker.zipcode()


def person() -> str:
    return faker.name()


def random_event_date() -> datetime.date:
    return faker.unique.date_this_century()


def random_plain_email() -> str:
    return faker.ascii_company_email()


for i in range(0,1000):
    document_1 = MailMerge(template_1)

    company = company_name()
    date = random_event_date()
    letter_date = '{:%d-%b-%Y}'.format(date)
    file_date = '{:%Y-%m-%d}'.format(date)
    document_1.merge(
        status=random_status(),
        phone_number="1-800-555-1212",
        company=company,
        purchases='$500,000',
        shipping_limit='$500',
        address=address(),
        date=letter_date,
        discount='5%',
        recipient=person())

    # Save the document as example 1
    document_1.write(f'files/{file_date}_{company}_letter.docx')