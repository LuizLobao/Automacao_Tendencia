from rich import print
from rich.layout import Layout

layout = Layout()

layout.split_column(
    Layout(name="upper"),
    Layout(name="lower")
)
layout["upper"].update(print('bom dia'))
print(layout)