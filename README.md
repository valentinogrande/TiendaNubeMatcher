# TIENDANUBEMATCHER

Sincroniza precios de productos desde un xlsx local contra una tienda NubeCommerce.

## Cómo funciona

1. **Lee productos** del archivo `products.xlsx` (todas las hojas)
2. **Obtiene productos** de la tienda via API de NubeCommerce
3. **Empareja** productos locales con los de la web usando un algoritmo de scoring por tokens, similitud de strings, marca, categoría, peso y precio
4. **Actualiza precios** en la web a `precio_local × 1.07`

## Productos del xlsx

El archivo `products.xlsx` debe tener una estructura con:

- Filas de **encabezado**: una celda con texto (≤4 palabras) que se toma como marca para los productos siguientes
- Filas de **producto**: texto descriptivo + precio numérico

Las columnas pueden variar por hoja; el parser detecta automáticamente valores de texto y precio.

## Configuración

Variables en `main.py`:

| Variable | Descripción |
|---|---|
| `ACCESS_TOKEN` | Token de API de NubeCommerce |
| `STORE_ID` | ID de la tienda |
| `USER_ID` | ID del usuario |
| `LOCAL_FILE` | Ruta al archivo xlsx (default: `products.xlsx`) |
| `AUTO_CONFIRM_SAFE` | Si `True`, confirma automáticamente cambios SAFE |
| `MIN_SCORE` | Puntaje mínimo para considerar un match (default: 40) |

## Matching

El algoritmo compara cada producto local contra productos web candidatos (misma categoría, o todos si no tiene categoría) y calcula un score basado en:

- Tokens compartidos (+10 cada uno, más para tokens importantes como variedades de queso/fiambre)
- Similitud de strings (hasta +50)
- Marca (+40 si coincide, -250 si difiere)
- Categoría (+30 si coincide, -150 si difiere)
- Peso (+20 si compatible, -120 si difiere)
- Precio (+15 si compatible, -60/-120 si difiere)

Estados: `SAFE`, `REVIEW`, `BLOCKED`.

## Uso

```bash
python3 main.py
```

El script pide confirmación antes de cada cambio de precio (excepto si `AUTO_CONFIRM_SAFE = True`).
