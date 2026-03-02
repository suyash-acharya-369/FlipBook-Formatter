Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "waitress-serve --port=5055 app:app", 0, False
WshShell.Run "ngrok http 5055", 0, False
