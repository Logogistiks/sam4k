import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget
from PyQt5.QtGui import QColor, QBrush
from PyQt5.QtCore import QTimer
import random

# Memory handler to store fetched data
class MemoryHandler:
    def __init__(self):
        self.data = []  # list of lists

    def append(self, new_row):
        self.data.append(new_row)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Live Table Example")

        self.memory = MemoryHandler()

        # Setup table widget
        self.table = QTableWidget()
        self.table.setColumnCount(5)  # Example column count
        self.table.setRowCount(0)

        layout = QVBoxLayout()
        layout.addWidget(self.table)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Timer to simulate serial fetch every 1 second
        self.timer = QTimer()
        self.timer.timeout.connect(self.fetch_and_update)
        self.timer.start(1000)

    def fetch_and_update(self):
        # Simulate fetching a row of 5 random numbers
        new_data = [random.randint(0, 100) for _ in range(5)]
        self.memory.append(new_data)

        # Insert a new row in the table
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)

        for col, value in enumerate(new_data):
            item = QTableWidgetItem(str(value))

            # Example formatting: color cell based on value
            if value > 70:
                item.setBackground(QBrush(QColor("lightgreen")))
            elif value < 30:
                item.setBackground(QBrush(QColor("lightcoral")))

            # Optional: bold font for high values
            if value > 90:
                font = item.font()
                font.setBold(True)
                item.setFont(font)

            # Insert item into table
            self.table.setItem(row_position, col, item)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
