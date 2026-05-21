# Project Overview

## Summary

An end-to-end compliance automation pipeline built 
to replace a fully manual monthly audit process 
across ~319 service locations for two corporate 
accounts.

A Power Automate Desktop robot scrapes a customer 
portal nightly, Python evaluates compliance against 
8 checks per location, and results are written to 
Excel for review and reporting.

---

## The Problem

Each service location must maintain active 
compliance documents at all times:

- Branch pest control business license
- Certificate of Insurance (COI)
- Technician pesticide license
- Technician client certification
- IPM / cGMP certification
- Annual facility assessment
- Quarterly pest trend report
- Quarterly pesticide usage log

With ~319 locations across two accounts, manually 
checking 8 documents per location every month is 
not sustainable. Documents expire, names are 
inconsistent, some belong to specific technicians, 
and some require current-year versions only.

This pipeline automates collection and evaluation 
of all compliance data monthly.

---

## Architecture
