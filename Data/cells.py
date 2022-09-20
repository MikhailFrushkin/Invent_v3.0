from peewee import *

from Data.connect_DB import *


class BaseModel(Model):
    class Meta:
        database = dbhandle


class Cells(BaseModel):
    id = PrimaryKeyField(null=False)
    place = CharField(max_length=25, verbose_name='Местоположение')
    code = CharField(max_length=8, verbose_name='Артикул')
    name = TextField(verbose_name='Описание товара')
    num = IntegerField(default=0, verbose_name='Физические запасы')
    num_dost = IntegerField(default=0, verbose_name='Доставка')
    num_sell = IntegerField(default=0, verbose_name='Продано')
    num_reserve = IntegerField(default=0, verbose_name='Резерв')
    num_free = IntegerField(default=0, verbose_name='Доступно')
    num_check = IntegerField(default=0, verbose_name='Посчитано')
    delta = IntegerField(default=0, verbose_name='Разница')

    @staticmethod
    def add_art(place, code, number=0, name='Лишний артикул на ячейке', num=0, num_reserve=0, num_dost=0, num_sell=0,
                num_free=0):
        row = Cells(
            place=place, code=code, name=name, num=num, num_dost=num_dost, num_sell=num_sell, num_reserve=num_reserve,
            num_free=num_free, num_check=number)
        row.save()

    class META:
        database = dbhandle
        db_table = 'Cells'
        order_by = ['place']


class Check(BaseModel):
    id = PrimaryKeyField(null=False)
    place = CharField(max_length=25)
    code = CharField(max_length=8)
    num = IntegerField(default=0)

    class META:
        database = dbhandle
        db_table = 'Check'
        order_by = ['place']
