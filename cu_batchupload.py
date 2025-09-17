import tkinter as tk
from tkinter import filedialog, messagebox
import metadata_utils as mu
import pyexcel
import requests
from fuzzywuzzy import fuzz

#---------------
#Helper Articles
#---------------

def preprocess_data():
    #-Get File
    input_file = input_entry.get()
    if not input_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return

    file = mu.read_csv(input_file)
    headers = file[0]

    #Get_inventory
    inventory = mu.read_csv("inventory.csv")

    #-Initial Filter
    file = [x for x in file if x[file[0].index("Include On Chicago Unbound")].lower() == "true" and not x[file[0].index("Chicago Unbound URL")]]

    #Next Filter
    ready_file = [headers]
    review_file = [headers]

    for line in file:
        flag = False

        # Question Mark check
        q_count = line[headers.index("Article Title")].count("?") if line[headers.index("Source Type")] != "Book" else line[headers.index("Book Title")].count("?")
        if q_count > 4:
            flag = True

        if not line[headers.index("Citation Year")]:
            flag = True

        # Fuzzy Check
        for entry in inventory:
        
            if all([line[headers.index("Generic Citation")], entry[1]]):
                if fuzz.ratio(line[headers.index("Generic Citation")], entry[1]) > 95:
                    flag = True
                    break
            elif not all([line[headers.index("Generic Citation")], entry[1]]):
                lawcites_title = line[headers.index("Article Title")] if line[headers.index("Source Type")] != "Book" else line[headers.index("Book Title")]
                if fuzz.ratio(lawcites_title, entry[0]) > 95:
                    flag = True
                    break
            else:
                flag = True
        if not flag:
            ready_file.append(line)
        else:
            review_file.append(line)

    output_location = output_entry.get()
    output_location = "/".join([x for x in output_entry.get().split("/")][:-1] + ["review.csv"])
    mu.write_csv(output_location, review_file)
    
    return ready_file

def isbad(url):
    headers = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}
    if "doi" in url.lower():
        try:
            doi = "/".join(url.split(".org")[1:]).strip("/")
            return not 200 <= requests.get(f"https://api.crossref.org/works/{doi}/", allow_redirects=True).status_code < 300
        except:
            return True
    else:
        try:
            response = requests.get(url, allow_redirects=True, headers=headers)
            return not 200 <= response.status_code < 300
        except requests.RequestException:
            return True

def select_input_file():
    filename = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
    input_entry.delete(0, tk.END)
    input_entry.insert(0, filename)

def select_output_file():
    filename = filedialog.asksaveasfilename(defaultextension=".xls", filetypes=[("Excel files", "*.xls")])
    output_entry.delete(0, tk.END)
    output_entry.insert(0, filename)

def get_include_index(input_filepath):
    # this function returns the index of the "Include On Chicago Unbound" heading, wherever it may lie in a given input file.
    # this helps prevent issues when someone strips out blank columns in a a file.
    headers = mu.read_csv(input_filepath)[0]
    return headers.index("Include On Chicago Unbound")

#---------------------
#The Convert Functions
#---------------------

def convert_book():
    # print("doing book!")
    input_file = input_entry.get()
    output_file = output_entry.get()
    
    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return

    #--------------------------------------------------

    #-Get Author Count for whole sheet

    try:
        ready_data = preprocess_data()
        max = 1
        for line in ready_data[1:]:
            author_count = line[5].lower().count("(auth)")
            if author_count > max:
                max = author_count

        #-Set Headings, with author fields on basis of max count
        data = [["title", "custom_citation", "publication_date", "publisher", "book_editors", "fulltext_url", "source_fulltext_url", "catalog_url", "document_type"]]
        for i in range(1, max + 1):
            data[0].extend([f"author{i}_fname", f"author{i}_lname"])

        #Initialize transformed line
        for line in ready_data[1:]:
            if line[get_include_index(input_file)].lower().strip() == "true":  #true test for "include in chicago unbound"
                #normal data
                new_line = [
                    line[9],
                    line[6],
                    line[21],
                    line[10],
                    "", #<--- will be "book-editors"
                    "", #<--- Will be fulltext_url
                    "", #<--- Will be source_fulltext_url
                    f"http://pi.lib.uchicago.edu/1001/cat/bib/{line[27] if line[27] else line[26]}" if any([line[27],line[26]]) else "", #<--- Will be catalog_url
                    "book"
                    ]


                #-Extract Editors
                if "(Ed)" in line[5]:
                    editors = [x.replace("(Ed)", "").strip() for x in line[5].split(",") if "(Ed)" in x]
                    if len(editors) == 1:
                        new_line[4] = editors[0]
                    elif len(editors) > 1:
                        new_line[4] = ", ".join(editors[:-1]) + " & " + editors[-1]

                #-Link work
                ext_url = []
                for field in [line[25],line[29],line[31],line[23]]:
                    if field is not None:
                        if " " in field.strip(" "):
                            link = field.strip(" ").split(" ")[0]
                        else:
                            link = field

                        #-if pdf, goes to fulltext_url
                        if ".pdf" in link.lower():
                            new_line[5] = link
                            break
                        else:
                            if "http" in link:
                                ext_url.append(link)

                #-sorting out external links for source_fulltext_url
                #-if one
                if len(ext_url) == 1:
                    if check_links.get():
                        if isbad(ext_url[0]) in [True, None]:
                            if messagebox.askyesno("Link Issue", f"Broken link found:\n---\n{ext_url[0]}\n---\nDo you want to fix the link(s)?"):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the new URL:", initialvalue=ext_url[0])
                                if input is not None:
                                    new_line[6] = input_url
                            else:
                                new_line[6] = ext_url[0]
                        else:
                            new_line[6] = ext_url[0]
                    else:
                        new_line[6] = ext_url[0]
                #-if multiple
                    #-checks them, filters on that basis
                    #-if not, checks if a proxy link, filters on that basis too
                elif len(ext_url) > 1:
                    if check_links.get():
                        new_links = [x for x in ext_url if not isbad(x)]
                        if len(new_links) == 1:
                            new_line[6] = new_links[0]
                        elif len(new_links) > 1:
                            if messagebox.askyesno("Link Issue", f"Multiple links found for :\n---\n{line[9]}\n---\nDo you want to manually fix the link? No link will be included if not."):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the new URL:")
                                if input is not None:
                                    new_line[6] = input_url
                        elif len(new_links) == 0:
                            if messagebox.askyesno("Link Issue", f"No workable links found for :\n---\n{line[9]}\n---\nDo you want to manually add the link? No link will be included if not."):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the URL:", initialvalue=ext_url[0])
                                if input is not None:
                                    new_line[6] = input_url
                    else:
                        filter_proxy = [x for x in ext_url if "proxy.uchicago" not in x]
                        if len(filter_proxy) == 1:
                            line[6] = filter_proxy[0]
                        else: #<--- this is the final resort when multipe links and no link checking turned on; for now just take first value
                            line[6] = ext_url[0]


                #author data
                if line[5].lower().count("(auth)") == 1:
                    if "(ed)" in line[5].lower():
                        for name in line[5].split(", "):
                            if "(auth)" in name.lower():
                                fname = name.split(" ")[0].strip()
                                lname = name.split(" ")[1].split("(")[0].strip()
                                new_line.extend([fname, lname])
                    else:
                        fname = line[5].split(" ")[0].strip()
                        lname = line[5].split(" ")[1].split("(")[0].strip()
                        new_line.extend([fname, lname])

                else:
                    for val in [x.strip() for x in line[5].split(", ") if "(Auth)" in x]:
                        # print(idx, val)
                        fname = val.split(" ")[0].strip()
                        lname = val.split(" ")[1].split("(")[0].strip()
                        new_line.extend([fname, lname])

                #add editors if no authors found
                if data[0].index("author1_lname") not in range(len(new_line)) and "(Ed)" in line[5]:
                    if line[5].lower().count("(ed)") == 1:
                        if "(ed)" in line[5].lower():
                            for name in line[5].split(", "):
                                if "(ed)" in name.lower():
                                    fname = name.split(" ")[0].strip()
                                    lname = name.split(" ")[1].split("(")[0].strip()
                                    new_line.extend([fname, lname])
                        else:
                            fname = line[5].split(" ")[0].strip()
                            lname = line[5].split(" ")[1].split("(")[0].strip()
                            new_line.extend([fname, lname])

                    else:
                        for val in [x.strip() for x in line[5].split(", ") if "(Ed)" in x]:
                            fname = val.split(" ")[0].strip()
                            lname = val.split(" ")[1].split("(")[0].strip()
                            new_line.extend([fname, lname])


                data.append(new_line)


        #write it!
        pyexcel.save_as(array=data, dest_file_name=output_file)
        messagebox.showinfo("Success", "File converted successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def convert_chapter():
    # print("doing chapter!")
    input_file = input_entry.get()
    output_file = output_entry.get()
    
    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return

    #--------------------------------------------------

    #-Get Author Count for whole sheet
    try:
        ready_data = preprocess_data()
        max = 1
        for line in ready_data[1:]:
            author_count = line[5].lower().count("(auth)")
            if author_count > max:
                max = author_count

        #-Set Headings, with author fields on basis of max count
        data = [["title", "custom_citation", "publication_date", "publisher", "book_editors", "fulltext_url", "source_fulltext_url", "catalog_url", "document_type", "source_publication", "catalog_button"]]
        for i in range(1, max + 1):
            data[0].extend([f"author{i}_fname", f"author{i}_lname"])

        #Initialize transformed line
        for line in ready_data[1:]:
            if line[get_include_index(input_file)].lower().strip() == "true":
                #normal data
                new_line = [
                    line[8],
                    line[6],
                    line[21],
                    line[10],
                    "", #<--- will be "book-editors"
                    "", #<--- Will be fulltext_url
                    "", #<--- Will be source_fulltext_url
                    '<p><a href="{x}" target="_blank">{x}</a></p>'.format(x = f"http://pi.lib.uchicago.edu/1001/cat/bib/{line[27] if line[27] else line[26]}") if any([line[27],line[26]]) else "", #<--- Will be catalog_url
                    "article",
                    line[9],
                    f"http://pi.lib.uchicago.edu/1001/cat/bib/{line[27] if line[27] else line[26]}" if any([line[27],line[26]]) else "", #<--- Will be catalog_button
                    ]

                
                #-Extract Editors
                if "(Ed)" in line[5]:
                    editors = [x.replace("(Ed)", "").strip() for x in line[5].split(",") if "(Ed)" in x]
                    if len(editors) == 1:
                        new_line[4] = editors[0]
                    elif len(editors) > 1:
                        new_line[4] = ", ".join(editors[:-1]) + " & " + editors[-1]

                #-Link work
                ext_url = []
                for field in [line[25],line[29],line[31],line[32]]:
                    if field is not None:
                        if " " in field.strip(" "):
                            link = field.strip(" ").split(" ")[0]
                        else:
                            link = field

                        #-if pdf, goes to fulltext_url
                        if ".pdf" in link.lower():
                            new_line[5] = link
                            break
                        else:
                            if "http" in link:
                                ext_url.append(link)

                #-sorting out external links for source_fulltext_url
                #-if one
                if len(ext_url) == 1:
                    if check_links.get():
                        if isbad(ext_url[0]) in [True, None]:
                            if messagebox.askyesno("Link Issue", f"Broken link found:\n---\n{ext_url[0]}\n---\nDo you want to fix the link(s)?"):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the new URL:", initialvalue=ext_url[0])
                                if input is not None:
                                    new_line[6] = input_url
                            else:
                                new_line[6] = ext_url[0]
                        else:
                            new_line[6] = ext_url[0]
                    else:
                        new_line[6] = ext_url[0]
                #-if multiple
                    #-checks them, filters on that basis
                    #-if not, checks if a proxy link, filters on that basis too
                elif len(ext_url) > 1:
                    if check_links.get():
                        new_links = [x for x in ext_url if not isbad(x)]
                        if len(new_links) == 1:
                            new_line[6] = new_links[0]
                        elif len(new_links) > 1:
                            if messagebox.askyesno("Link Issue", f"Multiple links found for :\n---\n{line[9]}\n---\nDo you want to manually fix the link? No link will be included if not."):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the new URL:")
                                if input is not None:
                                    new_line[6] = input_url
                        elif len(new_links) == 0:
                            if messagebox.askyesno("Link Issue", f"No workable links found for :\n---\n{line[9]}\n---\nDo you want to manually add the link? No link will be included if not."):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the URL:", initialvalue=ext_url[0])
                                if input is not None:
                                    new_line[6] = input_url
                    else:
                        filter_proxy = [x for x in ext_url if "proxy.uchicago" not in x]
                        if len(filter_proxy) == 1:
                            line[6] = filter_proxy[0]
                        else: #<--- this is the final resort when multipe links and no link checking turned on; for now just take first value
                            line[6] = ext_url[0]


                #author data
                if line[5].lower().count("(auth)") == 1:
                    if "(ed)" in line[5].lower():
                        for name in line[5].split(", "):
                            if "(auth)" in name.lower():
                                fname = name.split(" ")[0].strip()
                                lname = name.split(" ")[1].split("(")[0].strip()
                                new_line.extend([fname, lname])
                    else:
                        fname = line[5].split(" ")[0].strip()
                        lname = line[5].split(" ")[1].split("(")[0].strip()
                        new_line.extend([fname, lname])

                else:
                    for val in [x.strip() for x in line[5].split(", ") if "(Auth)" in x]:
                        # print(idx, val)
                        fname = val.split(" ")[0].strip()
                        lname = val.split(" ")[1].split("(")[0].strip()
                        new_line.extend([fname, lname])

                
                #add editors if no authors found
                if data[0].index("author1_lname") not in range(len(new_line)) and "(Ed)" in line[5]:
                    if line[5].lower().count("(ed)") == 1:
                        if "(ed)" in line[5].lower():
                            for name in line[5].split(", "):
                                if "(ed)" in name.lower():
                                    fname = name.split(" ")[0].strip()
                                    lname = name.split(" ")[1].split("(")[0].strip()
                                    new_line.extend([fname, lname])
                        else:
                            fname = line[5].split(" ")[0].strip()
                            lname = line[5].split(" ")[1].split("(")[0].strip()
                            new_line.extend([fname, lname])

                    else:
                        for val in [x.strip() for x in line[5].split(", ") if "(Ed)" in x]:
                            fname = val.split(" ")[0].strip()
                            lname = val.split(" ")[1].split("(")[0].strip()
                            new_line.extend([fname, lname])


                data.append(new_line)


        #write it!
        pyexcel.save_as(array=data, dest_file_name=output_file)
        messagebox.showinfo("Success", "File converted successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def convert_article():
    # print("doing article!")

    def isbad(url):
        try:
            status_code = requests.get(url, headers={'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/50.0.2661.102 Safari/537.36'}).status_code
            if status_code == 200:
                return False
            elif "doi.org" in url and status_code == 403:
                return False
        except:
            return True
        

    input_file = input_entry.get()
    output_file = output_entry.get()
    
    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select both input and output files.")
        return

    #--------------------------------------------------

    #-Get Author Count for whole sheet
    try:
        ready_data = preprocess_data()
        max = 1
        for line in ready_data[1:]:
            author_count = line[5].lower().count("(auth)")
            if author_count > max:
                max = author_count

        #-Set Headings, with author fields on basis of max count
        data = [["title", "custom_citation", "publication_date", "", "", "fulltext_url", "source_fulltext_url", "catalog_url", "document_type", "source_publication", "volnum", "issnum"]]
        for i in range(1, max + 1):
            data[0].extend([f"author{i}_fname", f"author{i}_lname"])

        #Initialize transformed line
        for line in ready_data[1:]:
            if line[get_include_index(input_file)].lower().strip() == "true":
                #normal data
                new_line = [
                    line[8],
                    line[6],
                    line[21],
                    "", #<--- was publishers
                    "", #<--- was be "book-editors"
                    "", #<--- Will be fulltext_url
                    "", #<--- Will be source_fulltext_url
                    f"http://pi.lib.uchicago.edu/1001/cat/bib/{line[27] if line[27] else line[26]}" if any([line[27],line[26]]) else "", #<--- Will be catalog_url
                    "article",
                    line[11],
                    line[12] if line[12] else "",
                    line[13] if line[13] else "",
                    #line[18] if line[18] else "" #<--- will be abstract
                    ]



                #-Extract Editors
                if "(Ed)" in line[5]:
                    editors = [x.replace("(Ed)", "").strip() for x in line[5].split(",") if "(Ed)" in x]
                    if len(editors) == 1:
                        new_line[4] = editors[0]
                    elif len(editors) > 1:
                        new_line[4] = ", ".join(editors[:-1]) + " & " + editors[-1]

                #-Link work
                ext_url = []
                for field in [line[25],line[29],line[31],line[32]]:
                    if field is not None:
                        if " " in field.strip(" "):
                            link = field.strip(" ").split(" ")[0]
                        else:
                            link = field

                        #-if pdf, goes to fulltext_url
                        if ".pdf" in link.lower():
                            new_line[5] = link
                            break
                        else:
                            if "http" in link:
                                ext_url.append(link)

                #-sorting out external links for source_fulltext_url
                #-if one
                if len(ext_url) == 1:
                    if check_links.get():
                        if isbad(ext_url[0]) in [True, None]:
                            if messagebox.askyesno("Link Issue", f"Broken link found:\n---\n{ext_url[0]}\n---\nDo you want to fix the link(s)?"):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the new URL:", initialvalue=ext_url[0])
                                if input is not None:
                                    new_line[6] = input_url
                            else:
                                new_line[6] = ext_url[0]
                        else:
                            new_line[6] = ext_url[0]
                    else:
                        new_line[6] = ext_url[0]
                #-if multiple
                    #-checks them, filters on that basis
                    #-if not, checks if a proxy link, filters on that basis too
                elif len(ext_url) > 1:
                    if check_links.get():
                        new_links = [x for x in ext_url if not isbad(x)]
                        if len(new_links) == 1:
                            new_line[6] = new_links[0]
                        elif len(new_links) > 1:
                            if len([x for x in new_links if "doi" in x]) == 1:
                                new_line[6] = [x for x in new_links if "doi" in x][0]
                            else:
                                if len([x for x in new_links if "ssrn" not in x]) == 1:
                                    new_line[6] = [x for x in new_links if "ssrn" not in x][0]
                                else:
                                    if messagebox.askyesno("Link Issue", f"Multiple links found for :\n---\n{line[8]}\n---\nDo you want to manually fix the link? No link will be included if not."):
                                        input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the new URL:")
                                        if input is not None:
                                            new_line[6] = input_url
                        elif len(new_links) == 0:
                            if messagebox.askyesno("Link Issue", f"No workable links found for :\n---\n{line[8]}\n---\nDo you want to manually add the link? No link will be included if not."):
                                input_url = tk.simpledialog.askstring(title="Fix Link", prompt="Enter the URL:", initialvalue=ext_url[0])
                                if input is not None:
                                    new_line[6] = input_url
                    else:
                        filter_proxy = [x for x in ext_url if "proxy.uchicago" not in x]
                        if len(filter_proxy) == 1:
                            line[6] = filter_proxy[0]
                        else: #<--- this is the final resort when multipe links and no link checking turned on; for now just take first value
                            line[6] = ext_url[0]


                #author data
                if line[5].lower().count("(auth)") == 1:
                    if "(ed)" in line[5].lower():
                        for name in line[5].split(", "):
                            if "(auth)" in name.lower():
                                fname = name.split(" ")[0].strip()
                                lname = name.split(" ")[1].split("(")[0].strip()
                                new_line.extend([fname, lname])
                    else:
                        fname = line[5].split(" ")[0].strip()
                        lname = line[5].split(" ")[1].split("(")[0].strip()
                        new_line.extend([fname, lname])

                else:
                    for val in [x.strip() for x in line[5].split(", ") if "(Auth)" in x]:
                        # print(idx, val)
                        fname = val.split(" ")[0].strip()
                        lname = val.split(" ")[1].split("(")[0].strip()
                        new_line.extend([fname, lname])


                #add editors if no authors found
                if data[0].index("author1_lname") not in range(len(new_line)) and "(Ed)" in line[5]:
                    if line[5].lower().count("(ed)") == 1:
                        if "(ed)" in line[5].lower():
                            for name in line[5].split(", "):
                                if "(ed)" in name.lower():
                                    fname = name.split(" ")[0].strip()
                                    lname = name.split(" ")[1].split("(")[0].strip()
                                    new_line.extend([fname, lname])
                        else:
                            fname = line[5].split(" ")[0].strip()
                            lname = line[5].split(" ")[1].split("(")[0].strip()
                            new_line.extend([fname, lname])

                    else:
                        for val in [x.strip() for x in line[5].split(", ") if "(Ed)" in x]:
                            fname = val.split(" ")[0].strip()
                            lname = val.split(" ")[1].split("(")[0].strip()
                            new_line.extend([fname, lname])


                data.append(new_line)


        #write it!
        pyexcel.save_as(array=data, dest_file_name=output_file)
        messagebox.showinfo("Success", "File converted successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

#---
#GUI
#---

materials_selector = {
                      "book" : convert_book,
                      "chapter" : convert_chapter,
                      "article" : convert_article
                      }

def on_radio_select():
    convert_button.config(command=materials_selector[material.get()])

# Create the main window
root = tk.Tk()
root.title("LawCites to BePress Converter")
root.geometry("350x375")

#Take of material selected
input_label = tk.Label(root, text=" ")
input_label.pack()
input_label = tk.Label(root, text="Select Material Type:")
input_label.pack()
material = tk.StringVar(value="article") #<--default value
book_button = tk.Radiobutton(root, text="Book", variable=material, value="book", command=on_radio_select)
book_button.pack()
book_chap_button = tk.Radiobutton(root, text="Book Chapter/section", variable=material, value="chapter", command=on_radio_select)
book_chap_button.pack()
article_button = tk.Radiobutton(root, text="Article", variable=material, value="article", command=on_radio_select)
article_button.pack()
input_label = tk.Label(root, text=" ")
input_label.pack()

# Input file selection
input_label = tk.Label(root, text="Input CSV file:")
input_label.pack()
input_entry = tk.Entry(root, width=50)
input_entry.pack()
input_button = tk.Button(root, text="Browse", command=select_input_file)
input_button.pack()

# Output file selection
output_label = tk.Label(root, text="Output XLS file:")
output_label.pack()
output_entry = tk.Entry(root, width=50)
output_entry.pack()
output_button = tk.Button(root, text="Browse", command=select_output_file)
output_button.pack()

# Verify links checkbox
input_label = tk.Label(root, text=" ")
input_label.pack()
check_links = tk.BooleanVar()
checkbutton = tk.Checkbutton(root, text="Verify Links?", variable=check_links)
checkbutton.pack(pady=(15,0))

# Convert button
convert_button = tk.Button(root, text="Convert", command=convert_article)
convert_button.pack()

# Start the GUI event loop
root.mainloop()