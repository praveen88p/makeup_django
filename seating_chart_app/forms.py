from django import forms

class SeatingChartForm(forms.Form):
    room_details_file = forms.FileField(label='Room Details Excel File', required=True)
    left_roll_numbers = forms.FileField(label='Left Roll Numbers Excel File', required=True)
    middle_roll_numbers = forms.FileField(label='Middle Roll Numbers Excel File', required=True)
    right_roll_numbers = forms.FileField(label='Right Roll Numbers Excel File', required=True)
