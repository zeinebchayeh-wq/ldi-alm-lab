"""
Sanity checks on the exported inputs.
A test is just a function starting with test_ that uses `assert`.
Run all tests with:  pytest
"""
from pathlib import Path

import pandas as pd

DATA = Path(__file__).resolve().parents[1] / "data"


def test_census_has_ten_participants():
    census = pd.read_csv(DATA / "census.csv")
    assert len(census) == 10


def test_census_weights_sum_to_one():
    census = pd.read_csv(DATA / "census.csv")
    assert abs(census["weight"].sum() - 1.0) < 1e-9


def test_mortality_covers_ages_0_to_110():
    mortality = pd.read_csv(DATA / "mortality.csv")
    assert mortality["age"].min() == 0
    assert mortality["age"].max() == 110


def test_mortality_value_matches_excel():
    # Known cell: Actuarial Assumption!B5 (male, age 0) = 0.00573
    mortality = pd.read_csv(DATA / "mortality.csv")
    assert abs(mortality.loc[0, "qx_male"] - 0.00573) < 1e-12
