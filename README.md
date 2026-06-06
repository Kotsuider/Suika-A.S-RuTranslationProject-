# Проект перевода Suika A.S+ на русский язык

## Общие сведения
Перевод делается gemini с последующей редактурой.
Патч предназначен для англофикатора Suika A.S+ 

## Статус

**Progress:**
`[██░░░░░░░░░░░░░░░░░] 15%`

- [x] Хук скриптов
- [x] Хук изображений
- [ ] Перевод системного текста
- [ ] Перевод графики
- [ ] Полный перевод
- [ ] Редактура

[Таблица с переводом](https://docs.google.com/spreadsheets/d/1uj5MKzgCKL6DsePytshkUoymAN5QBM8K/edit?usp=sharing&ouid=103224880279791937700&rtpof=true&sd=true) 


## Установка
1. Скачайте
2. Распакуйте в папку с игрой.
3. Профит!

## Редактирование патча
Достаём скрипты через [GARbro](https://github.com/morkt/GARbro/releases/tag/v1.5.44) из data01000.

Получаем распакованные папки, нам нужна script. 
Качаем xlsx.

```
python bgi_tool.py insert  scripts/ -x suika.xlsx -o patched/ --RU_F
```