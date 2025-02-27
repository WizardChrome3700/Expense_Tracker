#!/bin/bash

./expense_tracker/bin/activate
python ./Expense_loader.py
python ./Expense_calc.py