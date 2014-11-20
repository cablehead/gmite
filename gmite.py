import gdata.spreadsheet.service


class Gmite(object):
    def __init__(self, email, password, source='gmite', debug=False):
        self.conn = gdata.spreadsheet.service.SpreadsheetsService()
        self.conn.email = email
        self.conn.password = password
        self.conn.source = source
        self.conn.debug = debug
        self.conn.ProgrammaticLogin()

    def spreadsheet(self, key):
        return Gmite.Spreadsheet(self.conn, key)

    class Spreadsheet(object):
        def __init__(self, conn, key):
            self.conn = conn
            self.key = key

        def add_worksheet(self, title, col_count, row_count):
            entry = self.conn.AddWorksheet(
                title, row_count, col_count, self.key)
            return Gmite.Worksheet(self.conn, self.key, entry)

        def worksheet(self, title):
            match = [x for x in self.worksheets() if x.title == title]
            if match:
                return match[0]

        def worksheets(self):
            r = self.conn.GetWorksheetsFeed(key=self.key)
            return [
                Gmite.Worksheet(self.conn, self.key, entry)
                for entry in r.entry]

    class Worksheet(object):
        def __init__(self, conn, key, entry):
            self.conn = conn
            self.key = key
            self.entry = entry
            self.id = entry.id.text.split('/')[-1]
            self.col_count = int(entry.col_count.text)
            self.row_count = int(entry.row_count.text)
            self.title = entry.title.text

        def cells(self):
            r = self.conn.GetCellsFeed(self.key, self.id)
            return [
                Gmite.Cell(self.conn, self.key, self.id, entry)
                for entry in r.entry]

        def rows(self):
            r = self.conn.GetListFeed(self.key, self.id)
            return [
                Gmite.Row(self.conn, self.key, self.id, entry)
                for entry in r.entry]

        def update_cell(self, col, row, value):
            return self.conn.UpdateCell(row, col, value, self.key, self.id)

        def insert_row(self, row):
            entry = self.conn.InsertRow(row, self.key, self.id)
            return Gmite.Row(self.conn, self.key, self.id, entry)

        def delete(self):
            return self.conn.DeleteWorksheet(self.entry)

        def __repr__(self):
            return "<Worksheet: '%s'>" % self.title

    class Cell(object):
        def __init__(self, conn, key, wid, entry):
            self.conn = conn
            self.key = key
            self.wid = wid
            self.entry = entry
            self.id = entry.id.text.split('/')[-1]
            self.row = int(entry.cell.row)
            self.col = int(entry.cell.col)
            self.value = entry.cell.text

        def __repr__(self):
            return "<Cell: '%s: %s'>" % (self.id, self.value)

    class Row(object):
        def __init__(self, conn, key, wid, entry):
            self.conn = conn
            self.key = key
            self.wid = wid
            self.entry = entry
            self.id = entry.id.text.split('/')[-1]
            self.data = dict(
                (key, value.text) for key, value in entry.custom.items())

        def __repr__(self):
            return "<Row: '%s: %s'>" % (self.id, self.data)
