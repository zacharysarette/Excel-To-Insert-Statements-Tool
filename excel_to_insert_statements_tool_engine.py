import pylightxl as xl

def list_to_string(my_list):
    items_string = ""
    for item in my_list:
        items_string += "`" + item + "`, "
    return items_string.rstrip(", ")


def get_sheets(file_path):
    db = xl.readxl(fn=file_path)
    return db.ws_names


def run(file_name, sheet_name):
    db = xl.readxl(fn=file_name)
    fields = db.ws(ws=sheet_name).row(row=1)
    row_list = []
    for i, data in enumerate(db.ws(ws=sheet_name).rows):
        if i == 0:
            continue
        data_string = ""
        for d in data:
            data_string += "'" + str(d) + "', "
        row_list.append(
            "INSERT INTO `"
            + sheet_name
            + "` ("
            + list_to_string(fields)
            + ") VALUES ("
            + data_string.rstrip(", ")
            + ");"
        )
    output_string = ""
    for item in row_list:
        line = item + "\n"
        output_string += line

    return output_string
