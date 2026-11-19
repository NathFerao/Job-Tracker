"""
sample_data.py - fake but realistic emails so you can run and test the whole
pipeline WITHOUT connecting to Gmail (python main.py --demo).

They are written to cover the awkward cases on purpose:
  * different job sites (LinkedIn, Naukri, Indeed) and company career pages
  * confirmation emails that contain conditional sentences ("if selected...")
  * a company name that differs slightly between emails ("Pvt Ltd")
  * company only discoverable from the sender name
  * a job alert that must be ignored
  * the same job confirmed twice (LinkedIn + Naukri) -> must not create 2 rows
"""
from __future__ import annotations

from datetime import datetime

from models import RawEmail


def _mail(i, when, name, addr, subject, body) -> RawEmail:
    return RawEmail(
        id=f"demo{i:02d}", date=datetime.fromisoformat(when),
        sender_name=name, sender_email=addr, subject=subject, body=body.strip(),
    )


DEMO_EMAILS = [
    # 1. LinkedIn confirmation
    _mail(1, "2026-09-02T10:15:00", "LinkedIn", "jobs-noreply@linkedin.com",
          "Your application to Data Analyst at Zeta Analytics",
          """Your application was sent to Zeta Analytics

Data Analyst
Zeta Analytics · Bengaluru, Karnataka, India

Top skills for this role: Python, SQL, Tableau, Excel
https://www.linkedin.com/comm/jobs/view/3912345678?trackingId=abc123
"""),

    # 2. Naukri confirmation, company has "Pvt Ltd"
    _mail(2, "2026-09-05T18:40:00", "Naukri", "donotreply@naukri.com",
          "Application sent for Python Developer at Nimbus Software Pvt Ltd",
          """Hi, your application has been sent.

Location: Panaji, Goa
Skills: Python, Django, PostgreSQL, REST APIs, Git
https://www.naukri.com/job-listings-python-developer-nimbus-software-050926
"""),

    # 3. Indeed confirmation (role only in subject, company in body)
    _mail(3, "2026-09-10T09:05:00", "Indeed Apply", "indeedapply@indeed.com",
          "Indeed Application: Junior Backend Engineer",
          """The following items were sent to Coastal Labs. Good luck!

Junior Backend Engineer
Coastal Labs - Remote
Java, Spring Boot, MySQL, Docker
"""),

    # 4. Greenhouse confirmation with a conditional "interview" sentence (must stay 'Applied')
    _mail(4, "2026-09-12T14:22:00", "Orbit Robotics", "no-reply@us.greenhouse-mail.io",
          "Thank you for applying to Orbit Robotics",
          """Hi Aarav,

Thank you for applying for the Software Engineer – Platform position at Orbit Robotics.
We have received your application. If your qualifications match our needs, we will
contact you to schedule an interview. Unfortunately we cannot reply to every applicant.

Location: Pune, India (Hybrid)
You will work with Java, Kubernetes, AWS and Kafka.
"""),

    # 5. Lever confirmation
    _mail(5, "2026-09-15T11:00:00", "Finlytics", "no-reply@hire.lever.co",
          "Thanks for applying to Finlytics!",
          """Thanks for applying to the Data Engineer role at Finlytics.

Location: Mumbai, India
Tech stack: Spark, Airflow, SQL, Python and AWS.
"""),

    # 6. Same Nimbus job confirmed again by email from the company -> duplicate, ignore
    _mail(6, "2026-09-05T19:02:00", "Nimbus Careers", "careers@nimbussoft.in",
          "Application received - Python Developer at Nimbus Software",
          "Thank you for applying for the Python Developer position at Nimbus Software."),

    # 7. Nimbus interview invitation
    _mail(7, "2026-09-20T16:30:00", "Nimbus Careers", "careers@nimbussoft.in",
          "Interview Invitation – Python Developer at Nimbus Software",
          """Hello,

We would like to invite you to an interview for the Python Developer position.
Please let us know a convenient time on Friday.
"""),

    # 8. Orbit coding assessment (company only discoverable from the sender name)
    _mail(8, "2026-09-22T10:10:00", "Orbit Robotics Recruiting", "recruiting@orbitrobotics.com",
          "Complete your coding assessment",
          """As the next step in our hiring process, please complete the online assessment
on HackerRank within 5 days."""),

    # 9. Zeta rejection (via LinkedIn)
    _mail(9, "2026-09-24T08:45:00", "LinkedIn", "jobs-noreply@linkedin.com",
          "Your update from Zeta Analytics",
          """Thank you for your interest in the Data Analyst position at Zeta Analytics.
Unfortunately, we will not be moving forward with your application at this time."""),

    # 10. Orbit interview
    _mail(10, "2026-10-01T12:00:00", "Orbit Robotics Recruiting", "recruiting@orbitrobotics.com",
          "Next steps: Interview with Orbit Robotics",
          "We'd like to schedule an interview to discuss the platform engineering role."),

    # 11. Finlytics offer
    _mail(11, "2026-10-05T17:20:00", "Finlytics Talent Team", "talent@finlytics.io",
          "Offer of employment",
          """Dear candidate,

We are pleased to offer you the position of Data Engineer at Finlytics.
The offer letter is attached."""),

    # 12. Job alert -> must be ignored
    _mail(12, "2026-10-12T07:00:00", "LinkedIn Job Alerts", "jobalerts-noreply@linkedin.com",
          "Your job alert: Python Developer in Goa",
          "12 new jobs match your search. Apply to Data Analyst at Acme Corp and more."),

    # 13. Naukri confirmation, recent
    _mail(13, "2026-10-28T20:15:00", "Naukri", "donotreply@naukri.com",
          "Application sent for Machine Learning Engineer at Quantum Retail",
          """Your application has been sent.

Location: Remote
Skills: Python, TensorFlow, PyTorch, Machine Learning, Pandas
"""),

    # 14. Quantum viewed the application
    _mail(14, "2026-11-01T09:30:00", "Naukri", "donotreply@naukri.com",
          "Your application was viewed by Quantum Retail",
          "Quantum Retail viewed your application for Machine Learning Engineer."),

    # 15. LinkedIn confirmation, job in Ireland -> Outside India
    _mail(15, "2026-10-08T13:00:00", "LinkedIn", "jobs-noreply@linkedin.com",
          "Your application to Backend Engineer at Northwind Cloud",
          """Your application was sent to Northwind Cloud

Backend Engineer
Northwind Cloud · Dublin, Ireland

Go, Golang, Kubernetes, AWS, PostgreSQL
"""),

    # 16. Northwind interview (company career site, .ie domain)
    _mail(16, "2026-10-20T15:00:00", "Northwind Careers", "careers@northwindcloud.ie",
          "Interview Invitation – Backend Engineer at Northwind Cloud",
          "We would like to invite you to an interview for the Backend Engineer position."),

    # 17. LinkedIn confirmation, job in Switzerland -> Outside India
    _mail(17, "2026-11-03T09:00:00", "LinkedIn", "jobs-noreply@linkedin.com",
          "Your application to Cloud Engineer at Helvetia Systems",
          """Your application was sent to Helvetia Systems

Cloud Engineer
Helvetia Systems · Zurich, Switzerland

Docker, Kubernetes, Linux, CI/CD
"""),
]
