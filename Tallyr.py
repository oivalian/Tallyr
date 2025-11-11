import ttkbootstrap as ttk
from tkinter import filedialog, StringVar
from docx import Document as doc
import re
from os import path, getcwd


class Calls:
    def __init__(self):
        self.Filename = None
        self.File = None
        self.Count = 0
        self.Prompt = None
        self.Check_adv = False
        self.adverb_setter = []

    # count words in files from file_handler
    def word_counter(self, words):
        self.Count += len(words)
        count.set("Word Count: " + str(self.Count))
        counter.pack()

    # handle to files through Windows File Manager and feed them into counters
    def file_handler(self):
        self.Filename = filedialog.askopenfilenames(filetypes=[("Text Files", "*.txt;*.docx;*.doc")])

        try:
            for filepath in self.Filename:
                words = []      # store words in a list to pass into counter
                if filepath.endswith(".txt"):
                    with open(filepath, 'r') as self.File:
                        for lines in self.File:
                            line = lines.split()
                            words.extend(line)  # like .append(), adds line to words

                elif filepath.endswith(".docx") or filepath.endswith(".doc"):
                    self.File = doc(filepath)
                    for paras in self.File.paragraphs:
                        para = paras.text.split()
                        words.extend(para)      # like .append(), adds para to words
                else:
                    pass

                self.word_counter(words)    # pass words into counter
                if self.Check_adv:
                    self.adv_counter(words)     # if Check_adv selected, pass into adv_counter

        except Exception as e:
            self.Prompt = str(e)
            message.set(self.Prompt)
            error_window = ttk.Toplevel(root)
            error_window.geometry("500x350+600+600")
            error_window.resizable(False, False)
            error_frame = ttk.Frame(error_window)
            error_title = ttk.Label(error_window, text="An unexpected error occurred:", font="Helvetica 12")
            error_prompt = ttk.Label(error_frame, textvariable=message, font="Helvetica 9", wraplength=250)
            error_title.pack(pady="40", padx="50")
            error_prompt.pack(padx="10")
            error_frame.pack(padx="20")
            prompt.pack(pady="50")
            error_close = ttk.Button(error_window, text="OK", command=lambda: error_window.destroy())
            error_close.pack(pady="40", padx="20", ipadx="20", side="bottom")

    def reset(self):
        reset_window = ttk.Toplevel(root)
        reset_window.geometry("370x200+600+600")
        reset_window.resizable(False, False)
        reset_frame = ttk.Frame(reset_window)
        reset_prompt = ttk.Label(reset_window, text="Reset the counter?", font="Helvetica 12")
        reset_prompt.pack(pady="40", padx="50")
        reset_frame.pack(padx="50")

        reset_y = ttk.Button(reset_frame, text="OK", command=(lambda: (setattr(self, "Count", 0),
                                                                       reset_window.destroy(),
                                                                       count.set("Word Count: " + str(self.Count)),
                                                                       counter.pack())))
        reset_n = ttk.Button(reset_frame, text="CANCEL", command=(lambda: reset_window.destroy()))
        reset_y.pack(padx="20", ipadx="20", side="left")
        reset_n.pack(padx="20", side="right")


if __name__ == "__main__":
    current_dir = getcwd()
    icon_path = path.join(current_dir, "images\\icon.ico")
    open_file = path.join(current_dir, "images\\open_file.png")

    root = ttk.Window(themename="cyborg", title="Word Counter")
    root.iconbitmap(icon_path)
    root.title("countR")
    root.geometry("500x300+600+600")
    root.resizable(False, False)

    # VARS
    message = ttk.StringVar()
    count = ttk.IntVar()
    adv = ttk.BooleanVar()
    adverbs_total = ttk.StringVar()
    adverbs_list = ttk.StringVar()
    call = Calls()
    count_area = ttk.Frame(root)
    button_image = ttk.PhotoImage(file=open_file)
    button_image = button_image.subsample(2, 2)
    retrieve_button = ttk.Button(count_area, text=" Open File", image=button_image, compound="left",
                                 command=call.file_handler)
    reset_button = ttk.Button(count_area, text="Reset", command=call.reset)

    display_area = ttk.Frame(root)
    prompt = ttk.Label(display_area, textvariable=message, font="Helvetica 12")
    counter = ttk.Label(display_area, text="Word Count: 0", textvariable=count, font="Helvetica 16")

    count_area.pack(pady="50", padx="50")
    count.set("Word Count: 0")
    counter.pack()
    retrieve_button.pack(ipadx="50", pady="20", padx="10", side="left")
    reset_button.pack(pady="20", ipady="7", ipadx="25", padx="10", side="right")
    display_area.pack()

    root.mainloop()

