# Projekt Języki Skryptowe 

## Opis 
1.	Eksportujemy sprawozdanie w pliku Excela, sprawozdanie finansowe wybranej spółki notowanej na GPW( na przykładzie CD-Projekt).
2.	Z dostępnych danych program  wylicza ważniejsze wskaźniki finansowe (ROE, ROA, ROS,C/WK …  jest ich około 10) dla spółki, dla której pobraliśmy sprawozdanie finansowe.
3.	Przy użyciu biblioteki Matplotlib program rysuje zmiany ważniejszych wskaźników i ważniejszych kategorii finansowych( zysk netto, przychody ze sprzedaży, koszty, koszty działalności podstawowej).
4.	Tworzy dodatkowo wykres Du Ponta dla 2020 roku( często wykorzystywany do ogólnego spojrzenia na sytuację przedsiębiorstwa).

## Uruchomienie projektu 
W celu prawidłowego działania projektu należy pobrać następujące biblioteki:
-pandas;
-tkinter;
-matplotlib;
![image](https://user-images.githubusercontent.com/101069553/165181187-8a499dca-9046-4e4e-ad44-721b679c78ca.png)

### Pandas
Biblioteka pandas jest jednym z najbardziej rozbudowanych pakietów, do analizy danych w Python. Za jego pomocą możemy na przykład wczytywać, czyścić, modyfikować, a nawet analizować dane z Excela.
### Tkinter
Biblioteka tkinter umożliwia nam  tworzenie dość prostych programów okienkowych z wykorzystaniem w zasadzie wszystkich standardowych kontrolek systemowych.
### Matplotlib
Biblioteka matplotlib to bardzo obszerny i rozbudowany pakiet dający
niezmiernie dużo opcji wizualizacji danych. Umożliwia edycje i dostosywanie
wykresów.

## Przygotowanie danych 
Będziemy potrzebować sprawozdania finansowego wybranej spółki notowanej na GPW w pliku Excela.
![image](https://user-images.githubusercontent.com/101069553/165182752-eac36a38-5a35-455e-ba56-83cbcf7f4d4f.png)

Interesować nas będą takie dane jak:
-Zysk/(strata) netto"; 
-"AKTYWA OBROTOWE";
-"AKTYWA RAZEM";
-"Przychody ze sprzedaży";
-"KAPITAŁ WŁASNY";
-"AKTYWA OBROTOWE", "ZOBOWIĄZANIA KRÓTKOTERMINOWE";
-"Należności handlowe"; 
-"Koszty sprzedanych produktów, usług, towarów i materiałów";
-"Zapasy";
-"ZOBOWIĄZANIA DŁUGOTERMINOWE";

Istotnym warunkiem jest prawidłowy zapis wskaźników w programie Excel. Ponadto sprawozdanie finansowe ma obejmować statystyki z 4 lat.
![image](https://user-images.githubusercontent.com/101069553/165183082-f7b42e11-829b-48eb-9417-bcdcc287d484.png)
![image](https://user-images.githubusercontent.com/101069553/165183131-03b800c9-127c-415c-ad25-28cbb85273e3.png)
![image](https://user-images.githubusercontent.com/101069553/165183247-5569faba-259d-4abb-8cb7-90611256375a.png)
![image](https://user-images.githubusercontent.com/101069553/165183278-3648d21f-1fbc-4fc6-95e1-7eebc1b9b3a0.png)



 





