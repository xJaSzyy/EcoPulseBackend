# EcoPulse - Серверная часть
## Описание проекта
Backend-приложение для расчетов выбросов загрязняющих веществ от различных источников

## Запуск проекта
### Через Rider
1. Откройте решение (.sln файл) в Rider
2. Выберите проект запуска 
3. Запустите проект

### Через командную строку
Восстановление зависимостей
```
dotnet restore
```
Сборка проекта
```
dotnet build
```
Запуск в режиме разработки
```
dotnet run --project ./EcoPulseBackend/EcoPulseBackend.csproj
```
или из папки проекта:

```
cd ./EcoPulseBackend
dotnet run
```

## Swagger/OpenAPI
После запуска приложения документация API доступна по адресу:

* http://localhost:5000/swagger/index.html
