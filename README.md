# mme-excel

Będę pisał po polsku, bo nie chce mi się pisać po angielsku, chociaż piszę nie zawsze po polsku.

Kod jest bardzo spaghetti, ale jeśli go przytulić i otoczyć miłoscią to nawet będzie działać:

1. Przenieść plik Template.xlsx gdzieś (polecam Dokumenty)
2. Zmienić 19 linijkę pliku ThisAddIn.cs żeby była odpowiednia ścieżka do Template.
3. W 27 linijce wkleić swój własny klucz autoryzacyjny do Facebook Graph API, dostaniecie go stąd na przykład: https://www.facebook.com/login.php?next=https%3A%2F%2Fdevelopers.facebook.com%2Ftools%2Fexplorer
Każdy klucz jest ważny jakieś 1.5 godziny więc po tym czasie trzeba zdobyć nowy i wkleić go tam.
4. Uwaga, łatwo się wyraca.
5. Zmieniajce Template jeśli chcecie zmienić wygląd końcowego arkusza, do niego ładowane są statystyki z FB.
