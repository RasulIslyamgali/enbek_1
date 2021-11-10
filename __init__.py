from multiprocessing import Process
# from Sources.__main__ import colvir_enbek
# from Sources.send_outlook import info_outlook
from main import colvir_enbek


if __name__ == "__main__":
    p1 = Process(target=colvir_enbek)
    p1.start()
    # colvir_enbek() # for testing