import asyncio
import io
import os
import logging
import sys
from typing import Any, Dict
from aiogram import Bot, Dispatcher, Router, F, html
from aiogram.types import Message, KeyboardButton, ReplyKeyboardMarkup, ReplyKeyboardRemove, document, file
from aiogram.filters import CommandStart
from aiogram.filters.state import State, StatesGroup
import pandas as pd
import openpyxl
from aiogram.fsm.context import FSMContext
from dotenv import load_dotenv
import markups as nav

load_dotenv()
bot = Bot(os.getenv('TOKEN'))
dp = Dispatcher()
form_router = Router()

class Form(StatesGroup):
    group = State()
    find_group=State()
    inform=State()


@form_router.message(CommandStart())
async def command_start(message: Message, state: FSMContext) -> None:
    await message.answer("Привет! Я бот для работы с файлами Excel.")
    await message.answer ("Отправь мне файл, который необходимо обработать.")
 


@form_router.message(F.document)
async def file_downl(message: Message):
   global data
   try:
      file_id = message.document.file_id
      file = await bot.get_file(file_id)
      await message.answer ("Подождите, файл обрабатывается...")
      file_path=file.file_path
      my_object = io.BytesIO()
      MyBinaryIO = await bot.download_file(file_path, my_object)
      await message.answer ("Файл успешно загружен!",
            reply_markup=ReplyKeyboardMarkup(
               keyboard=[[KeyboardButton(text="Показать список групп")]],resize_keyboard=True,
            ),
      )
      data = pd.read_excel(MyBinaryIO)
   except Exception as e:
      await message.answer(f"Произошла ошибка при загрузке файла. {e}")
   return
 
@form_router.message(F.text =="Показать список групп")
async def all_groups(message: Message, state: FSMContext)-> None:
   groups = data['Группа'].unique()
   groups_str = ', '.join(groups)
   await message.answer(f'В файле содержатся данные о следующих группах: {groups_str}',
         reply_markup=ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="Выбрать группу")]],resize_keyboard=True,
         ),
   )

@form_router.message(F.text == "Выбрать группу")
async def report(message: Message, state: FSMContext) -> None:
   await state.set_state(Form.group)
   await message.answer("Введите номер группы: ")

@form_router.message(Form.group)
async def process_group(message: Message, state: FSMContext) -> None:
   global gruppa
   gruppa = message.text
   await state.update_data(gruppa=message.text)
   await state.set_state(Form.find_group)
   await message.answer(
      f"Произвести поиск по группе {html.quote(message.text)}?",
      reply_markup=ReplyKeyboardMarkup(
         keyboard=[
            [
               KeyboardButton(text="Да"),
               KeyboardButton(text="Нет"),
            ]
         ],
         resize_keyboard=True,
      ),
   )
   


@form_router.message(Form.find_group, F.text == "Нет")
async def process_no_group(message:Message, state: FSMContext) -> None:
   await state.update_data()
   await message.answer ("Выберите действие, нажав на кнопку",
         reply_markup=ReplyKeyboardMarkup(
            keyboard=[[KeyboardButton(text="Показать список групп"),KeyboardButton(text="Выбрать группу") ]],resize_keyboard=True,
         ),
   )


@form_router.message(Form.find_group, F.text == "Да")
async def process_find_group(message:Message, state: FSMContext) -> None:
   await state.set_state(Form.inform)
   groups = gruppa
   gr=len(data[data['Группа']==groups])
   if gr == 0:
      await message.answer(f'По группе с таким номером нет данных.')
   else:
      await message.reply(
         "Отлично!\nПодождите для загрузки отчета...",
         reply_markup=ReplyKeyboardRemove(),
      )
      groups = gruppa
      count_mark=data.shape[0]
      count_mark_gr=len(data[data['Группа']==groups])
      count_mark_group=data.loc[data['Группа']==groups, 'Личный номер студента'].nunique()
      count_id=data.loc[data['Группа']==groups, 'Личный номер студента'].unique()
      id_group=", ".join(map(str, count_id))
      form_contr=data.loc[data['Группа']==groups, 'Уровень контроля'].unique()
      forms=", ".join(form_contr)
      years=data.loc[data['Группа']==groups, 'Год'].unique()
      year=", ".join(map(str, years))
   
      await message.answer(f"""В исходном датасете содержалось {count_mark} оценок, из них {count_mark_gr} оценок относятся к группе {groups},
      \nВ датасете находятся оценки {count_mark_group} студентов со следующими личными номерами:  {id_group},
      \nИспользуемые формы контроля: {forms},
      \nДанные представлены по следующим учебным годам: {year}"""
      )


async def main():
   dp = Dispatcher()
   dp.include_router(form_router)
   await dp.start_polling(bot)

if __name__ == '__main__':
   logging.basicConfig(level=logging.INFO, stream=sys.stdout)
   asyncio.run(main())
   