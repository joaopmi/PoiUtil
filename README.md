# PoiUtil
Versão 2.0

Crie excel (XSSF | .xlsx) passando Strings como localizações das células. Ex: "A1","AF:B3"

Feito com base no Apache POI 3.17.
Alguns comentários podem estar desatualizados ou com ortografia errada devido às constantes alterações.
Foi criada apenas a variável para a fonte Calibri, pois é a usada no exemplo.

Exemplo no main da classe, apenas substitua o path em FileOutputStream ->
final FileOutputStream fos = new FileOutputStream(new File(path));
