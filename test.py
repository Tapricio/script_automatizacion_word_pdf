from tkinter import Tk
import main


root = Tk()
root.withdraw()

print("---menu---")
respuesta = input("eliga: ")
if respuesta == "1":
    main.main()
else:
    print("adios")