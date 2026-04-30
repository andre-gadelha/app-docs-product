from flask import Blueprint, render_template, redirect, url_for, session

main_bp = Blueprint('main', __name__)

@main_bp.route('/')
def index():
    # Simular usuário logado
    if 'user' not in session:
        session['user'] = 'Usuário Gemini'
    return render_template('index.html')

@main_bp.route('/logout')
def logout():
    session.pop('user', None)
    return redirect(url_for('main.index'))
