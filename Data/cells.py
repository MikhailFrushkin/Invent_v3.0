from peewee import *

from Data.Request.scrap_site import parse
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
    num_reserve = IntegerField(default=0, verbose_name='Резерв')
    num_free = IntegerField(default=0, verbose_name='Доступно')
    num_check = IntegerField(default=0, verbose_name='Посчитано')
    delta = IntegerField(default=0, verbose_name='Разница')
    box = IntegerField(default=1, verbose_name='Упаковок')

    @staticmethod
    def list():
        query = Cells.select()
        for row in query:
            print(row.id, row.place, row.code, row.name, row.num,
                  row.num_reserve, row.num_free, row.num_check, row.delta)
        return Cells.select()

    @staticmethod
    def add_art(place, art, number, name='Лишний артикул на ячейке', num=0, num_reserve=0, num_free=0, box=1):
        result_parse: dict = parse(art)
        if result_parse:
            name = result_parse['name']
            box = result_parse['box']
        row = Cells(
            place=place, code=art, name=name, num=num, num_reserve=num_reserve, num_free=num_free, num_check=number,
            box=box)
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

    @staticmethod
    def list():
        query = Cells.select()
        for row in query:
            print(row.id, row.place, row.code, row.name, row.num, row.num_free, row.num_reserve, row.updated_at)
        return Cells.select()

    class META:
        database = dbhandle
        db_table = 'Check'
        order_by = ['place']
