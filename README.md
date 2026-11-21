# Sync Connect Web

Cliente web Flutter de Sync Connect CRM. Este repositorio se limita al cliente web; el backend Node/Firebase y sus claves privadas permanecen en el VPS y no se incluyen aquí.

## Compilar

Requiere Flutter instalado y acceso a los paquetes de Pub.

```powershell
flutter pub get
flutter build web --release --dart-define=BOT_API_URL=https://syncconnect.online --dart-define=PUBLIC_WEB_URL=https://syncconnect.online
```

El resultado queda en `build/web`. El directorio `dist/` incluido en este repositorio ya contiene una compilación lista para publicar. Para alojarlo en Hostinger, publica el contenido de `dist/` en el directorio público del dominio raíz `syncconnect.online`. El backend debe responder desde `https://syncconnect.online`; los catálogos se publican en subdominios como `tienda.syncconnect.online`.

No se incluyen claves privadas, `.env`, datos de usuarios ni el servidor backend. La configuración pública de Firebase que necesita la aplicación web sí forma parte del cliente.
