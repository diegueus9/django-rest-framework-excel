# django-rest-framework-excel
A custome render to Excel 2010 xlsx files using openpyxl

Usage
-----

*views.py*

.. code-block:: python

    from rest_framework.views import APIView
    from rest_framework.settings import api_settings
    from rest_framework_excel.renderers import ExcelRenderer

    class MyView (APIView):
        renderer_classes = [ExcelRenderer] + api_settings.DEFAULT_RENDERER_CLASSES
        ...
