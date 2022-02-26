
def add_reg_text(paragraph, str):
    run = paragraph.add_run(str)

def add_bold_text(paragraph, str):
    run = paragraph.add_run(str)
    run.bold = True

def add_line_spaces(paragraph, num):
    run = paragraph.add_run()
    for i in range(0, num):
        run.add_break()

def bold_table(table):
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            paragraph = paragraphs[0]
            run_obj = paragraph.runs
            run = run_obj[0]
            run.bold= True

def bold_run(run):
    run.bold = True
