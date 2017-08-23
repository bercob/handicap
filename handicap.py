#!/usr/bin/python
# -*- coding: utf-8 -*-

import optparse, sys, webbrowser, os, time, signal
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import sqlite3

# constants; TODO: move them to json config
DB_PATH = "db/handicap.db"
FONTS_PATH = "fonts"
DEF_OUTPUT_PATH = "output/handicap.pdf"
DEF_DELIMITER = ";"
DEF_CHECK_FREQUENCY = 5
PLAYERS_TABLE_NAME = "players"
ROUNDS_TABLE_NAME = "rounds"
PLAYERS_COLS = ["id integer", "full_name text", "title text", "national_id text", "national_elo integer", "fide_elo integer", "birthdate date", "federation text", "sex text", "category text", 
"sk text", "club_id text", "club text", "fide_id text", "source text", "points text", "tb1 text", "tb2 text", "tb3 text", "tb4 text", "tb5 text", "ranking integer", "last_name text", 
"first_name text", "academic_title text"]
ROUNDS_COLS = ["round integer", "board integer", "white_national_id text", "black_national_id text", "white_player_id integer", "black_player_id integer", "white_result text", "black_result text", 
"loss_by_default text", "result text", "amount text", "white_res_rtg text", "black_res_rtg text"]

def signal_handler(signal, frame):
    sys.exit(0)

def help(parser):
	parser.print_help()

def parse_arguments(m_args):
	parser = optparse.OptionParser(usage = __file__ + " <exported_file_from_swissmanager>\n\n" + "handicap chess generator")
	parser.add_option("-o", "--output-path", default = DEF_OUTPUT_PATH, help = "default output pdf file (default is '%s')" % DEF_OUTPUT_PATH)
	parser.add_option("-d", "--delimiter", default = DEF_DELIMITER, help = "params delimiter in <exported_file_from_swissmanager> (default is '%s')" % DEF_DELIMITER)
	parser.add_option("-f", "--frequency", default = DEF_CHECK_FREQUENCY, help = "check frequency in sec (default is %d)" % DEF_CHECK_FREQUENCY)
	(options, args) = parser.parse_args()
	
	if m_args is not None:
		args = m_args
	
	if len(args) != 1:
		help(parser)
		sys.exit(1)

	return options, args

def init_styles():
	pdfmetrics.registerFont(TTFont("DejaVuSerif", "%s/DejaVuSerif.ttf" % FONTS_PATH))
	pdfmetrics.registerFont(TTFont("DejaVuSerif-Bold", "%s/DejaVuSerif-Bold.ttf" % FONTS_PATH))

	styles = getSampleStyleSheet()
	styles.add(ParagraphStyle(name="HNormal",
                               fontName="DejaVuSerif",
                               fontSize=12))
	
	styles.add(ParagraphStyle(name="HBold",
                               fontName="DejaVuSerif-Bold",
                               fontSize=12))
	return styles

def is_players_table(rows):
	return len(rows) > 0 and len(rows[0]) == len(PLAYERS_COLS)

def is_rounds_table(rows):
	return len(rows) > 0 and len(rows[0]) == len(ROUNDS_COLS)

def is_table_in_db(table_name):
	conn = get_connection()
	cursor = conn.cursor()
	cursor.execute("SELECT name FROM sqlite_master WHERE type = 'table' AND name = '%s'" % table_name)
	rows_count = len(cursor.fetchall())
	conn.close()
	return rows_count == 1

def get_table_name(rows):
	if is_players_table(rows):
		return PLAYERS_TABLE_NAME
	elif is_rounds_table(rows):
		return ROUNDS_TABLE_NAME
	else:
		print "unknown imported table"
		sys.exit(1)

def get_input_rows(filepath, delimiter):
	rows = []
	f_input = file(filepath,"r")
	for line in f_input:
		row = line.strip().split(delimiter)
		rows.append(row)
	f_input.close()
	return rows	

def get_connection():
	return sqlite3.connect(DB_PATH)

def get_cols_to_create(cols):
	return ",".join(["%s %s" % (s.split()[0], s.split()[1]) for s in cols])

def store_rows(rows):
	table_name = get_table_name(rows)

	conn = get_connection()
	conn.text_factory = lambda x: unicode(x, "utf-8", "ignore")
	c = conn.cursor()

	if not is_table_in_db(table_name):
		if is_players_table(rows):
			c.execute("CREATE TABLE %s(%s)" % (table_name, get_cols_to_create(PLAYERS_COLS)))
		elif is_rounds_table(rows):
			c.execute("CREATE TABLE %s(%s)" % (table_name, get_cols_to_create(ROUNDS_COLS)))
		else:
			print "unknown table to create"
			sys.exit(1)
	else:
		c.execute("DELETE FROM %s" % table_name)

	rows.pop(0)
	
	c.executemany("INSERT INTO %s VALUES (%s)" % (table_name, ",".join(["?" for s in rows[0]])), rows)
	
	conn.commit()

	conn.close()

def get_stored_rows(table_name):
	conn = get_connection()
	c = conn.cursor()
	select = get_select(table_name)
	if select:
		c.execute(select)
		rows = c.fetchall()
		conn.close()
		return rows
	else:
		return []

def get_round():
	conn = get_connection()
	cursor = conn.cursor()
	cursor.execute("SELECT max(round) FROM %s" % ROUNDS_TABLE_NAME)
	round = cursor.fetchone()[0]
	conn.close()
	return round

def get_select(table_name):
	if table_name == PLAYERS_TABLE_NAME:
		return "SELECT ranking, id, full_name, points, birthdate FROM %s ORDER BY ranking ASC, id ASC" % PLAYERS_TABLE_NAME
	elif table_name == ROUNDS_TABLE_NAME:
		if is_table_in_db(PLAYERS_TABLE_NAME):
			return """SELECT r.board, p_white.full_name, p_black.full_name 
			FROM %s r, %s p_white, %s p_black 
			WHERE p_white.id = r.white_player_id 
			AND p_black.id = r.black_player_id 
			AND r.round = (SELECT max(round) FROM %s)
			ORDER BY r.board ASC
			""" % (ROUNDS_TABLE_NAME, PLAYERS_TABLE_NAME, PLAYERS_TABLE_NAME, ROUNDS_TABLE_NAME)
		else:
			print "export the players"
			return ""
	else:
		print "unknown table to select"
		sys.exit(1)

def build_pdf(table_name, output_path):
	styles = init_styles()
	
	doc = SimpleDocTemplate(output_path, pagesize = landscape(A4),
                        rightMargin = 72,leftMargin = 72,
                        topMargin = 72,bottomMargin = 18)
	story = []
	
	stored_rows = get_stored_rows(table_name)

	if table_name == PLAYERS_TABLE_NAME:
		story.append(Paragraph(table_name, styles["HBold"]))
	elif table_name == ROUNDS_TABLE_NAME:
		if stored_rows:
			story.append(Paragraph("round %s" % get_round(), styles["HBold"]))
		else:
			story.append(Paragraph("Export the players, please.", styles["HBold"]))
	
	if stored_rows:
		table = Table(stored_rows)
		
		table.setStyle(TableStyle([
			("FONTNAME", (0, 0), (-1, -1), "DejaVuSerif"),
			("FONTSIZE", (0, 0), (-1, -1), 12),
			("ALIGN",(0,0),(-1,-1),"LEFT")
			]))
		story.append(table)
	
	doc.build(story)		

def open_pdf(output_path):
	webbrowser.open(r"file:///" + os.getcwd() + "/" + output_path)

def get_mtime(file_path):
	return os.stat(file_path)[8]

def main(m_args=None):
	(options, args) = parse_arguments(m_args)

	input_file_path = args[0]

	mtime_last = 0
	error = False

	while True:
		try:
			mtime_new = get_mtime(input_file_path)
			if mtime_new == mtime_last and not error:
				time.sleep(options.frequency)
				continue
			else:
				mtime_last = mtime_new
				error = False
			
			rows = get_input_rows(args[0], options.delimiter)
			
			store_rows(rows)
			
			build_pdf(get_table_name(rows), options.output_path)
			
			open_pdf(options.output_path)
			break #for debuging
			time.sleep(options.frequency)
		except (IOError, OSError), (errno, strerror):
			print "Error(%s): %s" % (errno, strerror)
			error = True
			time.sleep(options.frequency)

#-------------------------------    
if __name__ == "__main__":
	signal.signal(signal.SIGINT, signal_handler)
	main()
