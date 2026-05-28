"""Генератор CSV-таблицы для слияния наградных документов.

Скрипт создает файл awards_data.csv с 330 строками. Разделитель полей - точка
с запятой, кодировка - UTF-8 with BOM, поэтому файл корректно открывается в
Microsoft Excel без потери русских символов.
"""

from __future__ import annotations

import csv
import random
from pathlib import Path
from typing import Dict, List

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_CSV = BASE_DIR / "awards_data.csv"
ROW_COUNT = 330

HEADERS = [
    "document_id", "document_type", "title_text", "award_text", "recipient_name", "recipient_role",
    "child_name", "teacher_name", "place", "nomination", "age_group", "work_title",
    "contest_name", "organizer", "city", "issue_date", "signature_role", "signatory_name", "note"
]

LAST_NAMES = [
    "Иванов", "Петров", "Сидоров", "Кузнецов", "Смирнов", "Попов", "Васильев", "Соколов",
    "Михайлов", "Новиков", "Федоров", "Морозов", "Волков", "Алексеев", "Лебедев", "Семенов",
    "Егоров", "Павлов", "Козлов", "Степанов", "Николаев", "Орлов", "Андреев", "Макаров",
    "Никитин", "Захаров", "Зайцев", "Соловьев", "Борисов", "Яковлев"
]
FIRST_NAMES_M = [
    "Александр", "Михаил", "Даниил", "Артем", "Иван", "Максим", "Кирилл", "Егор", "Матвей", "Никита",
    "Тимофей", "Роман", "Дмитрий", "Владислав", "Павел"
]
FIRST_NAMES_F = [
    "Анна", "Мария", "София", "Алиса", "Виктория", "Екатерина", "Полина", "Дарья", "Ксения", "Арина",
    "Варвара", "Елизавета", "Анастасия", "Ульяна", "Вероника"
]
PATRONYMICS_M = ["Александрович", "Михайлович", "Сергеевич", "Иванович", "Дмитриевич", "Андреевич"]
PATRONYMICS_F = ["Александровна", "Михайловна", "Сергеевна", "Ивановна", "Дмитриевна", "Андреевна"]
TEACHER_FIRST = ["Елена", "Ольга", "Наталья", "Ирина", "Татьяна", "Светлана", "Марина", "Анна", "Юлия", "Галина"]
TEACHER_LAST = ["Соколова", "Петрова", "Иванова", "Кузнецова", "Васильева", "Новикова", "Федорова", "Морозова", "Орлова", "Макарова"]
TEACHER_PAT = ["Викторовна", "Александровна", "Сергеевна", "Игоревна", "Павловна", "Дмитриевна"]
NOMINATIONS = [
    "Изобразительное искусство", "Литературное творчество", "Декоративно-прикладное искусство",
    "Фотография", "Компьютерная графика", "Экологический плакат", "Социальный проект"
]
AGE_GROUPS = ["7-9 лет", "10-12 лет", "13-15 лет", "16-17 лет"]
WORKS = [
    "Мой город будущего", "Берег детства", "Семейная история", "Планета добрых дел",
    "Свет северного города", "Мы выбираем мир", "Окно в природу", "Герои рядом",
    "Традиции моей семьи", "Путешествие мечты", "Цифровой мир глазами детей", "Сохраним планету вместе"
]


def make_child_name(index: int) -> str:
    """Возвращает ФИО участника."""
    if index % 2 == 0:
        return f"{random.choice(LAST_NAMES)} {random.choice(FIRST_NAMES_M)} {random.choice(PATRONYMICS_M)}"
    return f"{random.choice(LAST_NAMES)}а {random.choice(FIRST_NAMES_F)} {random.choice(PATRONYMICS_F)}"


def make_teacher_name() -> str:
    """Возвращает ФИО педагога."""
    return f"{random.choice(TEACHER_LAST)} {random.choice(TEACHER_FIRST)} {random.choice(TEACHER_PAT)}"


def build_rows() -> List[Dict[str, str]]:
    """Формирует строки таблицы: дипломы, сертификаты и благодарственные письма."""
    random.seed(42)
    rows: List[Dict[str, str]] = []
    for index in range(1, ROW_COUNT + 1):
        child_name = make_child_name(index)
        teacher_name = make_teacher_name()
        nomination = NOMINATIONS[(index - 1) % len(NOMINATIONS)]
        age_group = AGE_GROUPS[(index - 1) % len(AGE_GROUPS)]
        work_title = WORKS[(index - 1) % len(WORKS)]

        if index % 11 == 0:
            document_type = "Благодарственное письмо"
            title_text = "БЛАГОДАРСТВЕННОЕ ПИСЬМО"
            award_text = "выражается благодарность"
            recipient_name = teacher_name
            recipient_role = "педагогу-наставнику"
            place = "—"
            note = f"за подготовку участника конкурса: {child_name}"
        elif index % 5 == 0:
            place = ["I место", "II место", "III место"][(index // 5) % 3]
            document_type = "Диплом"
            title_text = f"ДИПЛОМ {place.upper()}"
            award_text = "награждается"
            recipient_name = child_name
            recipient_role = "призер конкурса"
            note = "за высокий результат и творческий подход к выполнению конкурсной работы"
        else:
            document_type = "Сертификат участника"
            title_text = "СЕРТИФИКАТ УЧАСТНИКА"
            award_text = "подтверждает участие"
            recipient_name = child_name
            recipient_role = "участник конкурса"
            place = "участие"
            note = "за участие во всероссийском конкурсе творческих работ детей"

        rows.append({
            "document_id": f"VKT-2026-{index:04d}",
            "document_type": document_type,
            "title_text": title_text,
            "award_text": award_text,
            "recipient_name": recipient_name,
            "recipient_role": recipient_role,
            "child_name": child_name,
            "teacher_name": teacher_name,
            "place": place,
            "nomination": nomination,
            "age_group": age_group,
            "work_title": work_title,
            "contest_name": "Всероссийский конкурс творческих работ детей «Творчество без границ»",
            "organizer": "Оргкомитет всероссийского конкурса творческих работ детей",
            "city": "Москва",
            "issue_date": "25.05.2026",
            "signature_role": "Председатель оргкомитета",
            "signatory_name": "Романова Е. В.",
            "note": note,
        })
    return rows


def save_csv(rows: List[Dict[str, str]], output_path: Path) -> None:
    """Сохраняет строки в CSV-файл с разделителем ';'."""
    with output_path.open("w", encoding="utf-8-sig", newline="") as file_obj:
        writer = csv.DictWriter(file_obj, fieldnames=HEADERS, delimiter=";")
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    """Создает CSV-файл для последующего слияния в Word."""
    rows = build_rows()
    save_csv(rows, OUTPUT_CSV)
    print(f"Создан файл: {OUTPUT_CSV}")
    print(f"Количество строк данных: {len(rows)}")


if __name__ == "__main__":
    main()
