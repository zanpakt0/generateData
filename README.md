# Data Generation Program
Сonsole program that generated humans data

- В файле names.xls хранится база имен, разделенная по половому признаку;
- в файле manSurname.xls хранится база мужских фамилий;
- В файле womanSurname.xls хранится база женских фамилий;
- в файле manPatronymic.xls хранится база мужских отчеств;
- В файле womanPatronymic.xls хранится база женских отчеств;
- В файле country.xls хранится база стран;
- В файле region.xls хранится база областей России;
- В файле street.xls хранится база улиц (взял выборку только Москвы, но можно подставить туда что угодно);
- Индекс берется случайным образом (6 знаков, от 100000 до 999999);
- Номер дома допускается от 1 до 300;
- Номер квартиры допускается от 1 до 500;
- Дата рождения берется с 1970 года и до текущего времени;

Для каждого поля класса Human написан getter и setter.
- getter возвращает текущее значение поля у экземпляра.
- setter присваивает случайное значение данному полю, в рамках ограничений.

Класс Generator - генерирует случайные данные. Возвращает массив структур Human.

Класс RighterToExcel, предназначен для записи подготовленных данных в Excel файл.
Имеет методы righter и createRow.
- righter - непосредственно производитзапись в файл
- createRow - создает строку и записывает ее на лист.
