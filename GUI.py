from tkinter import *
from tkinter import filedialog
from tkinter import ttk

windown = Tk()
windown.geometry("1150x500")
windown.title("CHẤM THI WORD TỰ ĐỘNG - TRƯỜNG DHSP HÀ NỘI")
label = Label(windown, text="Trường đại học sư phạm Hà Nội",  font=('Arial', 12))
label.pack(padx= 20, pady= 20)
label.place(x=1, y=1)
label2 = Label(windown,text="Khoa công nghệ thông tin", font=('Arial', 9))
label2.place(x=1, y=25)

label3 = Label(windown,text="TẢI LÊN BÀI LÀM", font=('Arial', 12))
label3.place(x=100, y=90)

def open_file():
    filepath = filedialog.askopenfilename(initialdir="D:\\KhoaLuanTotNghiep",  title="Chon bai thi")
    label_add_file = Label(windown, text=filepath)
    label_add_file.place(x=100, y=180)


btn_upload = Button(windown, text="Tải lên", fg='blue', font=('Arial', 9), command=open_file)
btn_upload.pack()
btn_upload.place(x=130, y= 120)

accept_file = Label(windown,text="File được chấp nhận:", font=('Arial', 7))
accept_file.place(x=30, y=150)
file_type = Label(windown, text=".docx", fg="red", font=('Arial', 7))
file_type.place(x=120, y=150)
file_path = Label(windown, text="File tải lên:"  , font=('Arial', 9))
file_path.place(x=30, y=180)



btn_score = Button(windown, text="Chấm điểm", font=('Arial', 9))
btn_score.place(x=120, y= 450)

fill_data = Label(windown, text="ĐIỀN DỮ LIỆU CHẤM THI", font=('Arial', 12))
fill_data.place(x=700, y=70)


# fillsizepaper
file_size_paper = Label(windown, text="Kích thước khổ giấy: ", font=('Arial', 10))
file_size_paper.place(x= 550, y= 100)

paper_h = Label(windown, text="Height: ", font=('Arial', 10))
paper_h.place(x= 700, y= 100)
entry_h = Entry(windown).place(x=750, y= 100, width=50)
labelinch1 = Label(windown, text="inch", fg="red", font=('Arial', 10)).place(x= 800, y=100)
paper_w = Label(windown, text="Width:", font=('Arial', 10))
paper_w.place(x=850, y= 100)
entry_w = Entry(windown).place(x=900, y= 100, width=50)
labelinch2 = Label(windown, text="inch", fg="red", font=('Arial', 10)).place(x= 950, y=100)


# fillmarginpaper
file_margin_paper = Label(windown, text="Căn lề khổ giấy: ", font=('Arial', 10)).place(x= 550, y= 130)

txt_top = Label(windown, text="Top:", font=('Arial', 10)).place(x= 650, y= 130)
E_margin_top = Entry(windown).place(x=680, y= 130, width=50)

txt_left = Label(windown, text="Left:", font=('Arial', 10)).place(x= 750, y= 130)
E_margin_left = Entry(windown).place(x=780, y= 130, width=50)

txt_right = Label(windown, text="Right:", font=('Arial', 10)).place(x= 850, y= 130)
E_margin_right = Entry(windown).place(x=890, y= 130, width=50)

txt_bottom = Label(windown, text="Bottom:", font=('Arial', 10)).place(x= 950, y= 130)
E_margin_bottom = Entry(windown).place(x=1000, y= 130, width=50)

# SelectFontname
txt_font_name = Label(windown, text="Font:", font=('Arial', 10)).place(x= 550, y= 160)
E_font_name = Entry(windown).place(x=610, y= 160)

# SelectFontSize
txt_font_size = Label(windown, text="Size:", font=('Arial', 10)).place(x=750, y= 160)
E_font_size = Entry(windown).place(x=800, y=160, width=70)
txt_Pt = Label(windown, text="Pt", fg="red", font=('Arial', 10)).place(x=870, y= 160)

# Margin
txt_margin = Label(windown, text="Căn lề:", font=('Arial', 10)).place(x=550, y=190)
Combo_type = ttk.Combobox(windown, values=["Center(1)", "Right(2)", "Justify(3)"]).place(x= 610, y=190)

# HEADING
txt_heading = Label(windown, text="HEADING", font=('Arial', 10)).place(x=550, y=220)
txt_doc_apply = Label(windown,text="Các đoạn văn bản cần áp dụng", font=('Arial', 10)). place(x=550, y= 250)
e_head1 = Entry(windown).place(x= 550, y= 280, width= 270)
e_head2 = Entry(windown).place(x= 830, y= 280, width= 270)
e_head3 = Entry(windown).place(x= 550, y= 310, width= 270)




windown.mainloop()
