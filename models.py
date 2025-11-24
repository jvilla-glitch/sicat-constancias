# -*- coding: utf-8 -*-
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from werkzeug.security import generate_password_hash, check_password_hash
from flask_login import UserMixin

db = SQLAlchemy()

class Constancia(db.Model):
    __tablename__ = 'constancias'
    
    id = db.Column(db.Integer, primary_key=True)
    
    # Campos que se mostrarán en el documento
    fecha = db.Column(db.String(50), nullable=False)
    folio = db.Column(db.String(50), nullable=False, unique=True)
    vigencia = db.Column(db.String(50), nullable=False)
    expediente = db.Column(db.String(50), nullable=False)
    nombre_propietario = db.Column(db.String(200), nullable=False)
    domicilio_propietario = db.Column(db.Text, nullable=False)
    giro = db.Column(db.String(200), nullable=False)
    denominado = db.Column(db.String(200), nullable=False)
    ubicado = db.Column(db.Text, nullable=False)
    entre_calles = db.Column(db.String(200), nullable=False)
    colonia = db.Column(db.String(100), nullable=False)
    ciudad = db.Column(db.String(100), nullable=False)
    codigo_postal = db.Column(db.String(10), nullable=False)
    nombre_comisionado = db.Column(db.String(200), nullable=False)
    
    # Campos internos (NO se mostrarán en la impresión)
    recibo_pago = db.Column(db.String(100))
    referencia_comprobante = db.Column(db.String(100))
    constancias_avala = db.Column(db.Text)
    
    # Metadatos
    fecha_creacion = db.Column(db.DateTime, default=datetime.utcnow)
    fecha_modificacion = db.Column(db.DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
    
    def __repr__(self):
        return f'<Constancia {self.folio}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'fecha': self.fecha,
            'folio': self.folio,
            'vigencia': self.vigencia,
            'expediente': self.expediente,
            'nombre_propietario': self.nombre_propietario,
            'domicilio_propietario': self.domicilio_propietario,
            'giro': self.giro,
            'denominado': self.denominado,
            'ubicado': self.ubicado,
            'entre_calles': self.entre_calles,
            'colonia': self.colonia,
            'ciudad': self.ciudad,
            'codigo_postal': self.codigo_postal,
            'nombre_comisionado': self.nombre_comisionado,
            'recibo_pago': self.recibo_pago,
            'referencia_comprobante': self.referencia_comprobante,
            'constancias_avala': self.constancias_avala,
            'fecha_creacion': self.fecha_creacion.strftime('%Y-%m-%d %H:%M:%S') if self.fecha_creacion else None
        }


class Usuario(UserMixin, db.Model):
    __tablename__ = 'usuarios'
    
    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(200), nullable=False, unique=True)
    password_hash = db.Column(db.String(200), nullable=False)
    activo = db.Column(db.Boolean, default=True)
    fecha_creacion = db.Column(db.DateTime, default=datetime.utcnow)
    ultimo_acceso = db.Column(db.DateTime)
    
    def set_password(self, password):
        """Establece la contraseña hasheada"""
        self.password_hash = generate_password_hash(password)
    
    def check_password(self, password):
        """Verifica la contraseña"""
        return check_password_hash(self.password_hash, password)
    
    def __repr__(self):
        return f'<Usuario {self.nombre}>'
    
    def to_dict(self):
        return {
            'id': self.id,
            'nombre': self.nombre,
            'activo': self.activo,
            'fecha_creacion': self.fecha_creacion.strftime('%Y-%m-%d %H:%M:%S') if self.fecha_creacion else None,
            'ultimo_acceso': self.ultimo_acceso.strftime('%Y-%m-%d %H:%M:%S') if self.ultimo_acceso else None
        }






