# !/usr/bin/env python3

import peewee

from peewee import *
from playhouse.shortcuts import RetryOperationalError


class MyMySQL(RetryOperationalError, MySQLDatabase):
    pass


config = {'host': '127.0.0.1', 'password': '024464', 'port': 3306, 'user': 'wangjianfeng', 'charset': 'utf8mb4'}
database = MyMySQL('edu', **config)


class BaseModel(Model):
    class Meta:
        database = database


class Question(BaseModel):
    id = BigIntegerField(null=False)  # id 题目id在数据库中是自增
    class_name = CharField()  # 类型名称，待用（数理化分类）
    problem_type = CharField()  # 问题类型 待用：选择题、填空题等等
    problem = CharField(null=False)  # 问题内容，保存问题的主体
    answer = CharField()  # 问题的答案，待用

    class Meta:
        db_table = "paper_info_question"
