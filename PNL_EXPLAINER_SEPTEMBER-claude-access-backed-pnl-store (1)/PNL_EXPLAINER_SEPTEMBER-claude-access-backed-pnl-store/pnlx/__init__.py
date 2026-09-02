"""PnL Explainer - fixed-income PnL attribution in pure Python.

A migration of the Excel/VBA PNL Explainer to Python.  The whole model runs on
numpy column arrays: no worksheet, no formula strings, no Bloomberg add-in.

Only three third-party packages are used, each for one job:

    numpy      the numeric core
    PyYAML     the configuration file
    openpyxl   the Excel deliverable

Everything else is the standard library.

Layout
------
    config       typed configuration, loaded from YAML
    daycount     day-count conventions, coupon schedules, accrued interest
    curves       zero curves and the interpolation the model is calibrated on
    instruments  position dataclasses and their column-array "books"
    loaders      CSV readers that build the books
    frameworks   spread-framework resolution and the attribution chains
    bondmath     price/yield, duration, convexity, pull to par
    engine       the bond / futures / swap attribution engines
    aggregate    portfolio bridge, per-asset-class comparison, data quality
    snail        the snail-trail methodology
    report       CSV, JSON and Excel outputs
    cli          the entry point
"""

__version__ = "2.0.0"
__all__ = ["__version__"]
