from flask_wtf import FlaskForm
from wtforms import StringField
from wtforms.validators import DataRequired


class MyForm(FlaskForm):
    url = StringField('Input', validators=[DataRequired()])
    excecao = StringField('Exceção', validators=[DataRequired()])
    output = StringField('Output', validators=[DataRequired()])