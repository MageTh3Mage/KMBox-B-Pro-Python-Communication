# KMBox B Pro – Python Communication Library

A minimal Python library for communicating with the KMBox B Pro over a CH340 based COM port connection.

---

## ⚙️ Features

* Automatic detection of CH340 device
* Send mouse movement and click commands
* Debug output

---

## 🧪 Commands

| Method           | Description                   |
| ---------------- | ----------------------------- |
| `move(x, y)`     | Move mouse by (x, y) pixels   |
| `left_click()`   | Simulate a left mouse click   |
| `right_click()`  | Simulate a right mouse click  |
| `middle_click()` | Simulate a middle mouse click |
| `close()`        | Close the serial connection   |
| `is_connected()` | Returns `True` if connected   |
