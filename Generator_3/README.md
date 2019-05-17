# Excel_macros

## Program
Projekt stworzony przy użyciu:
* Microsoft Excel i VBA: Office 365 *(Poprawione, aby było zgodne z wersją 2010)*

## Założenia

Cele założone przy tworzeniu projektu:
1. Sprawy nie mogą być z tego samego dnia.
2. Jeżeli jednostka ma mniej niż 4 dni, to liczba spraw jest równa liczbie dni. Maksymalna liczba spraw to 3.
3. Jeśli istnieje, to przynajmniej jedna sprawa musi być zrealizowana przez jednostkę.
4. Ostateczny wynik sortowany po korzeniu i następnie po jednostce.

## Jak używać?
Aby uruchomić makro należy wejść w plik, do którego chcemy wykonać Generator_3.

W arkuszu uruchamiamy na początku makro "randomize_table", następnie po zakończeniu działania uruchamiamy makro "generator". Makro będzie działało około 5 minut. Po zakończeniu działania w arkuszu utworzonym przez "generator" włączamy makro "preparing_report".

Te działania zostawią dwa arkusze: wejściowy i wyjściowy z danymi określonymi w projekcie.

## Sposób realizacji

### Dla wszystkich:
1. Przed rozpoczęciem działania wyłączane jest odświeżanie ekranu i obliczanie ustawione jest na manualne.
2. Wszystkie zmienne są deklarowane. W przypadku, gdy nie było pewności dot. możliwej wielkości zmiennej wybierana była większa zmienna. Nie wszystkie zmienne zostały wykorzystane, ale ze względu na możliwość rozwijania tego kodu pozostały nieusunięte.
3. Po zakończeniu działanie włączane jest odświażanie ekranu i obliczanie ustawione jest na automatyczne.

### Makro "randomize_table":
1. W pierwszej wolnej kolumnie dla każdego niepustego wiersza (od drugiego) wypełnia pole różnymi wartościami na podstawie funkcji RND.
2. Sortuje kolejność zgodnie z liczbami wypełnionymi w punkcie 1.
3. Usuwa kolumnę dodaną w punkcie 1.

### Makro "generator":
1. Sprawdza liczbę rekordów.
2. Nadaje nazwy tworzonym arkuszom na podstawie aktualnej daty i godziny.
3. Dane z arkusza głównego kopiuje do pomocniczego i usuwa w nim niepotrzebne kolumny, zmienia format daty na "rrrr-mm-dd", sortuje zgodnie z zadaną kolejnością - data, korzeń, jednostka.
4. Ponowne przeklejenie danych do nowego arkusza, wypełnienie jedynie wartościami.
5. Dodanie nazw kolejnych kolumn wykorzystywanych w trakcie tworzenia makra (zbędne, ale zostawione w kodzie na potrzeby debugowania)
6. Wypełnienie kolumn z datami na podstawie drugiego arkusza, używając dwóch pętli do while.
7. Pierwsza część makra działa dla osób z liczbą dat mniejszą od 4. Porównuje datę sprawy z kolejnymi datami pracy danej jednostki.
8. Dla osób z liczbą dat mniejszą od 4 wypisuje dla każdej 1 sprawę, a następnie dla każdej jednostki drugą i ewentualnie trzecią.

Pierwsza kolumna:

9. Dla 1 daty i 0 liczby zamkniętych spraw przez daną osobę wyszukuje dowolną sprawę danej osoby i koloruje komórkę na czerwono (i wszystkie kolejne sytuacje, w których dana osoba nie zamknęła sprawy będą oznaczone kolorem czerwonym)
10. Dla 1 daty i > 0 liczby zamkniętych spraw szuka sprawy zamkniętej przez daną osobę.
11. Dla 2 albo 3 dat i 0 zamkniętych spraw wyszukuje dowolną sprawę danej osoby.
12. Dla 2 albo 3 dat i > 0 zamkniętych spraw, to szuka sprawę ze statusem "TAK". Jeśli nie znajdzie to szuka dowolnej sprawy danej osoby (zgodnej z datą).

Po zakończeniu rozpoczyna pracę z drugą kolumną:

13. Patrz pkt. 11
14. Dla 2 dat i > 0 zamkniętych i czerwonej pierwszej sprawy szuka sprawy ze statusem "TAK" na drugą kolumnę. Jeśli nie znajdzie takiej sprawy uzupełnia dowolną inną (zgodnej z datą).

Po zakończeniu rozpoczyna pracę z trzecią kolumną:

15. Dla 3 dat i 0 zamkniętych wyszukuje dowolną sprawę danej osoby.
16. Dla 3 dat i > 0 zamkniętych i obie kolumny czerwone szuka sprawy ze statusem "TAK". Jeśli nie znajdzie takiej sprawy uzupełnia dowolną inną (zgodnej z datą)
17. Jeśli jakaś osoba ma > 0 zamkniętych, ale wszystki kolumny są czerwone to wyszukuje pierwszą dostępną sprawę ze statusem "TAK" i zamienia ją ze sprawą o tej samej dacie z innym statusem.

Druga część makra, dla osób z > 3 datami.

18. Deklaracja zmiennych, nieużywanych w innej części makra.
19. Obecnie przeszukuje tylko pierwszą kolumnę, ale uzupełnia również pozostałe dwie.
20. Zmienne ustawione tak, aby uporządkować "sektory działania", aby możliwie rozsiać daty. Pierwsza data jest z przedziału od 1 do liczby dni dzielonej przez 3 (część całkowita). Druga data jest od liczby o jeden większej do liczby dwa razy większej od ostatniej liczby z sektora pierwszego. Trzecia data jest z przedziału od jeden większego od zakończenia poprzedniego sektora do końca.
21. Dla pierwszych 2 spraw wyszukuje dowolnych spraw (zgodnych z wylosowaną datą w sektorze) i jeśli status jest <> TAK to zaznacza kolorem czerwonym.
22. Jeśli obie sprawy są czerwone to szuka = TAK dla 3 daty. Jeśli nie znajdzie to wrzuca czerwoną.
23. Po wpisaniu wszystkich spraw sprawdza, czy osoba na pewno ma chociaż jedną o innym kolorze niż czerwony, jeśli nie to szuka pierwszej sprawy ze statusem TAK i zamienią ją ze sprawą w odpowiednim sektorze.
24. Na koniec usuwa arkusz roboczy, wcześniej wyłączając powiadomienia, a po usunięciu ponownie je włączając.

### Makro "preparing_report":
1. Deklaruje zmienne podobne do tych w "generator".
2. Zaznacza docelowy arkusz do oddania.
3. Usuwa z niego zbędne kolory (czerwony i biały).
4. Ustawia obramowania czarną nieprzerywaną linią z każdej strony, bez linii wewnętrznych. Kolor czarny bez innych efektów.
5. Nagłówkowi nadaje bordowy kolor tła, pogrubioną białą czcionkę.
6. Usuwa zbędne kolumny i dopasowuje te, które zostają.
7. Wyłącza podświetlenie lini i zaznacza A1 dla estetyki.

## Potencjalne problemy
* Większa liczba jednostek niż 499
* Większa liczba rekordów niż 100 000
* Problem mogą (ale nie muszą) stanowić inne ustawienia systemowe daty.
* Przy "pechowym" rozlosowaniu makro może działać zdecydowanie dłużej. Przy testach makro najdłużej działało ok. 15 minut. Losowanie radykalnie zmniejszyło ten czas.