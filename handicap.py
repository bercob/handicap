#!/usr/bin/python
# -*- coding: utf-8 -*-

import optparse, sys, webbrowser, os, time, signal, logging
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import SimpleDocTemplate, Paragraph, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import sqlite3
import xlsxwriter

# config; TODO: move them out
VERSION = "1.1.0"
AUTHOR = "Balogh Peter <bercob@gmail.com>"
DEF_SM_EXPORTED_FILE_PATH = "sm_exported_files/Exp.TXT"
DB_PATH = "db/handicap.db"
FONTS_PATH = "fonts"
DEF_OUTPUT_PATH = "output/handicap.xlsx"
ALLOWED_OUTPUT_FORMATS = ["pdf", "xlsx"]
DEF_HANDICAPS_CONFIG_PATH = "conf/handicaps.csv"
LOG_FILE_PATH = "log/handicap.log"
DEF_DELIMITER = ";"
DEF_CHECK_FREQUENCY = 5
PLAYERS_TABLE_NAME = "players"
ROUNDS_TABLE_NAME = "rounds"
HANDICAPS_TABLE_NAME = "handicaps"
PLAYERS_COLS = ["id integer", "full_name text", "title text", "national_id text", "national_rating integer", "fide_rating integer", "birthdate date", "federation text", "sex text", "category text", 
"sk text", "club_id text", "club text", "fide_id text", "source text", "points text", "tb1 text", "tb2 text", "tb3 text", "tb4 text", "tb5 text", "ranking integer", "last_name text", 
"first_name text", "academic_title text"]
ROUNDS_COLS = ["round integer", "board integer", "white_national_id text", "black_national_id text", "white_player_id integer", "black_player_id integer", "white_result text", "black_result text", 
"loss_by_default text", "result text", "amount text", "white_res_rtg text", "black_res_rtg text"]
HANDICAP_COLS = ["worst_player_time text", "better_player_time text", "diff_rating_from integer", "diff_rating_to integer", "better_player_rating_from integer", "better_player_rating_to integer"]

# translation
SK = { "t_id" : "číslo", "t_board" : "šachovnica", "t_full_name" : "hráč", "t_national_rating" : "národné elo", "t_white_handicap" : "čas bieleho", "t_result" : "výsledok", 
	"t_black_handicap" : "čas čierneho", "t_fide_rating" : "fide elo", "t_ranking" : "poradie", "t_points" : "body", "t_birthdate" : "dátum narodenia", "t_rating" : "elo", 
	"t_white" : "biely", "t_black" : "čierny", "t_number" : "číslo", "handicap" : "čas", "color" : "farba figúr", "t_round" : "kolo", "t_players" : "hráči"}
LANG = SK
# end of config

def signal_handler(signal, frame):
	sys.exit(0)

def help(parser):
	parser.print_help()

def set_logging():
	logging.basicConfig(filename = LOG_FILE_PATH, level = logging.INFO, format = "%(asctime)s %(message)s")
	sh = logging.StreamHandler(sys.stdout)
	sh.setFormatter(logging.Formatter("%(asctime)s %(message)s"))
	logging.getLogger().addHandler(sh)

def parse_arguments(m_args):
	parser = optparse.OptionParser(usage = "%s\n\nhandicap chess generator\n\nauthor: %s" % (__file__, AUTHOR))
	parser.add_option("-e", "--exported-file-path", default = DEF_SM_EXPORTED_FILE_PATH, help = "swissmanager exported file path (default is '%s')" % DEF_SM_EXPORTED_FILE_PATH)
	parser.add_option("-o", "--output-path", default = DEF_OUTPUT_PATH, help = "default output file (default is '%s')" % DEF_OUTPUT_PATH)
	parser.add_option("-c", "--handicaps-config-path", default = DEF_HANDICAPS_CONFIG_PATH, help = "default handicaps config file (default is '%s')" % DEF_HANDICAPS_CONFIG_PATH)
	parser.add_option("-d", "--delimiter", default = DEF_DELIMITER, help = "params delimiter in <exported_file_from_swissmanager> (default is '%s')" % DEF_DELIMITER)
	parser.add_option("-f", "--frequency", default = DEF_CHECK_FREQUENCY, help = "check frequency in sec (default is %d)" % DEF_CHECK_FREQUENCY)
	parser.add_option("-n", "--national-rating", action="store_true", help = "calculate handicap based on national rating (else fide rating)", default = False)
	parser.add_option("-p", "--classic-pairing", action="store_true", help = "generate classic pairing table", default = False)
	parser.add_option("-t", "--with-timestamp", action="store_true", help = "generate output with timestamp", default = False)
	parser.add_option("-v", "--version", action="store_true", help = "get version", default = False)
	(options, args) = parser.parse_args()
	
	if m_args is not None:
		args = m_args

	if options.version:
		print VERSION
		sys.exit(0)
	
	return options, args

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
		logging.error("unknown imported table")
		sys.exit(1)

def get_input_rows(filepath, delimiter):
	rows = []
	f_input = file(filepath,"r")
	for line in f_input:
		row = line.strip().split(delimiter)
		rows.append(row)
	f_input.close()
	return rows	

def get_handicaps(filepath, delimiter):
	rows = get_input_rows(filepath, delimiter)
	store_rows(rows, HANDICAPS_TABLE_NAME)

def get_connection():
	return sqlite3.connect(DB_PATH)

def get_cols_to_create(cols):
	return ",".join(["%s %s" % (s.split()[0], s.split()[1]) for s in cols])

def store_rows(rows, table_name):
	conn = get_connection()
	conn.text_factory = lambda x: unicode(x, "utf-8", "ignore")
	c = conn.cursor()

	if not is_table_in_db(table_name):
		if is_players_table(rows):
			c.execute("CREATE TABLE %s(%s)" % (table_name, get_cols_to_create(PLAYERS_COLS)))
		elif is_rounds_table(rows):
			c.execute("CREATE TABLE %s(%s)" % (table_name, get_cols_to_create(ROUNDS_COLS)))
		elif table_name == HANDICAPS_TABLE_NAME:
			c.execute("CREATE TABLE %s(%s)" % (table_name, get_cols_to_create(HANDICAP_COLS)))
		else:
			logging.error("unknown table to create")
			sys.exit(1)
	else:
		c.execute("DELETE FROM %s" % table_name)

	if len(rows) > 0:
		rows.pop(0)
	
	rowcount = 0
	if len(rows) > 0:
		c.executemany("INSERT INTO %s VALUES (%s)" % (table_name, ",".join(["?" for s in rows[0]])), rows)
		conn.commit()
		rowcount = c.rowcount

	conn.close()

	return rowcount

def get_stored_rows(table_name, options, header = False):
	conn = get_connection()
	cursor = conn.cursor()
	select = get_select(table_name, options)
	if select:
		cursor.execute(select)
		rows = cursor.fetchall()
		if header:
			rows.insert(0, [description[0] for description in cursor.description])
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

def get_select(table_name, options):
	if options.national_rating:
		rating = 'national_rating'
	else:
		rating = 'fide_rating'
			
	if table_name == PLAYERS_TABLE_NAME:
		select_dict = { 'rating' : rating, 'players_table_name' : PLAYERS_TABLE_NAME }
		return "SELECT ranking '%(t_ranking)s', full_name '%(t_full_name)s', points '%(t_points)s', %(rating)s '%(t_rating)s' FROM %(players_table_name)s ORDER BY ranking ASC, id ASC" % dict(select_dict.items() + LANG.items())
	elif table_name == ROUNDS_TABLE_NAME:
		if is_table_in_db(PLAYERS_TABLE_NAME):
			select_dict = { 'rating' : rating, 'handicaps_table_name' : HANDICAPS_TABLE_NAME, 'rounds_table_name' : ROUNDS_TABLE_NAME, 'players_table_name' : PLAYERS_TABLE_NAME }
			if options.classic_pairing:
				return """SELECT r.board %(t_board)s, 
							(SELECT 
								CASE WHEN p_white.%(rating)s > p_black.%(rating)s THEN h.better_player_time ELSE h.worst_player_time END 
								FROM %(handicaps_table_name)s h 
								WHERE abs(p_white.%(rating)s - p_black.%(rating)s) BETWEEN h.diff_rating_from AND h.diff_rating_to AND 
									((p_white.%(rating)s > p_black.%(rating)s AND p_white.%(rating)s BETWEEN h.better_player_rating_from AND h.better_player_rating_to)
										OR
									(p_black.%(rating)s >= p_white.%(rating)s AND p_black.%(rating)s BETWEEN h.better_player_rating_from AND h.better_player_rating_to))
								LIMIT 1
							) '%(t_white_handicap)s',
							p_white.full_name %(t_full_name)s, p_white.%(rating)s %(t_rating)s, p_white.points %(t_points)s,
							(SELECT 
								CASE WHEN p_black.%(rating)s > p_white.%(rating)s THEN h.better_player_time ELSE h.worst_player_time END 
								FROM %(handicaps_table_name)s h 
								WHERE abs(p_white.%(rating)s - p_black.%(rating)s) BETWEEN h.diff_rating_from AND h.diff_rating_to AND
									((p_white.%(rating)s > p_black.%(rating)s AND p_white.%(rating)s BETWEEN h.better_player_rating_from AND h.better_player_rating_to)
										OR
									(p_black.%(rating)s >= p_white.%(rating)s AND p_black.%(rating)s BETWEEN h.better_player_rating_from AND h.better_player_rating_to))
								LIMIT 1
							) '%(t_black_handicap)s',
							p_black.full_name '%(t_full_name)s', p_black.%(rating)s '%(t_rating)s', p_white.points %(t_points)s
							FROM %(rounds_table_name)s r, %(players_table_name)s p_white, %(players_table_name)s p_black 
							WHERE p_white.id = r.white_player_id 
							AND p_black.id = r.black_player_id
							AND r.round = (SELECT max(round) FROM %(rounds_table_name)s)
							ORDER BY r.board ASC
						""" % dict(select_dict.items() + LANG.items())
			else:
				return """SELECT number '%(t_number)s', full_name '%(t_full_name)s', board '%(t_board)s', 
							(SELECT 
								CASE WHEN %(rating)s > opponent_rating THEN h.better_player_time ELSE h.worst_player_time END 
								FROM handicaps h 
								WHERE abs(%(rating)s - opponent_rating) BETWEEN h.diff_rating_from AND h.diff_rating_to AND 
									((%(rating)s > opponent_rating AND %(rating)s BETWEEN h.better_player_rating_from AND h.better_player_rating_to)
										OR
									(opponent_rating >= %(rating)s AND opponent_rating BETWEEN h.better_player_rating_from AND h.better_player_rating_to))
								LIMIT 1
							) '%(handicap)s',
							color '%(color)s', %(rating)s '%(t_rating)s'
						FROM (
							SELECT (SELECT count(*) FROM players ps WHERE p.full_name >= ps.full_name) number, 
								p.full_name, 
								r.board,
								CASE WHEN r.white_player_id = p.id THEN '%(t_white)s' ELSE '%(t_black)s' END color,
								p.%(rating)s,
								(SELECT p_opponent.%(rating)s FROM players p_opponent WHERE (r.white_player_id = p_opponent.id OR r.black_player_id = p_opponent.id) AND p_opponent.id != p.id) opponent_rating
							FROM players p, rounds r
							WHERE r.round = (SELECT max(round) FROM rounds)
							AND (r.white_player_id = p.id OR r.black_player_id = p.id)
						) ORDER BY number ASC """ % dict(select_dict.items() + LANG.items())

		else:
			logging.warning("export the players")
			return ""
	elif table_name == HANDICAPS_TABLE_NAME:
		return "SELECT * FROM %s ORDER BY diff_rating_from ASC" % HANDICAPS_TABLE_NAME
	else:
		logging.error("unknown table to select")
		sys.exit(1)

def get_output_path(options):
	if options.with_timestamp:
		filename, file_extension = os.path.splitext(options.output_path)
		return "%s.%s%s" % (filename, time.strftime("%Y%m%d-%H%M%S"), file_extension)
	else:
		return options.output_path

def get_output_path_extension(output_path):
	filename, file_extension = os.path.splitext(output_path)
	return file_extension[1:]

def build_output(table_name, output_path, options):
	stored_rows = get_stored_rows(table_name, options, True)

	output_format = get_output_path_extension(output_path)

	if output_format == 'pdf':
		build_pdf(table_name, output_path, stored_rows)
	elif output_format == 'xlsx':
		build_xlsx(table_name, output_path, stored_rows)
	else:
		logging.error("unknown output format %s; allowed formats: %s" % (output_format, ",".join(ALLOWED_OUTPUT_FORMATS)))
		sys.exit(1)

def build_pdf(table_name, output_path, stored_rows):
	styles = init_pdf_styles()
	
	doc = SimpleDocTemplate(output_path, pagesize = landscape(A4),
                        rightMargin = 72,leftMargin = 72,
                        topMargin = 72,bottomMargin = 18)
	story = []
	
	if table_name == PLAYERS_TABLE_NAME:
		story.append(Paragraph("%(t_players)s" % dict(LANG.items()), styles["HBold"]))
	elif table_name == ROUNDS_TABLE_NAME:
		if stored_rows:
			vars = { 'round': get_round() }
			story.append(Paragraph("%(t_round)s %(round)s" % dict(LANG.items() + vars.items()), styles["HBold"]))
		else:
			story.append(Paragraph("Export the players, please.", styles["HBold"]))
	
	if stored_rows:
		table = Table(stored_rows)
		
		table.setStyle(TableStyle([
			("FONTNAME", (0, 0), (-1, 0), "DejaVuSerif-Bold"),
			("FONTNAME", (0, 1), (-1, -1), "DejaVuSerif"),
			("FONTSIZE", (0, 0), (-1, -1), 12),
			("ALIGN",(0,0),(-1,-1),"LEFT")
			]))
		story.append(table)
	
	doc.build(story)		

def init_pdf_styles():
	pdfmetrics.registerFont(TTFont("DejaVuSerif", "%s/DejaVuSerif.ttf" % FONTS_PATH))
	pdfmetrics.registerFont(TTFont("DejaVuSerif-Bold", "%s/DejaVuSerif-Bold.ttf" % FONTS_PATH))

	styles = getSampleStyleSheet()
	styles.add(ParagraphStyle(name = "HNormal",
                               fontName = "DejaVuSerif",
                               fontSize = 12,
                               spaceAfter = 10))
	
	styles.add(ParagraphStyle(name = "HBold",
                               fontName = "DejaVuSerif-Bold",
                               fontSize = 12,
                               spaceAfter = 10))
	return styles

def build_xlsx(table_name, output_path, stored_rows):
	workbook = xlsxwriter.Workbook(output_path)
	worksheet = workbook.add_worksheet()

	for row, cols in enumerate(stored_rows):
		for col, data in enumerate(cols):
			worksheet.write(row, col, data)
	
	worksheet.set_column(0, col, 20)
	
	workbook.close()

def open_output(output_path):
	webbrowser.open(r"file:///" + os.getcwd() + "/" + output_path)

def get_mtime(file_path):
	return os.stat(file_path)[8]

def main(m_args=None):
	set_logging()

	logging.info("starting")
	
	reload(sys)
	sys.setdefaultencoding("utf8")

	(options, args) = parse_arguments(m_args)

	mtime_last = 0
	error = False

	while True:
		try:
			mtime_new = get_mtime(options.exported_file_path)
			if mtime_new == mtime_last and not error:
				time.sleep(options.frequency)
				continue
			else:
				mtime_last = mtime_new
				error = False
			
			get_handicaps(options.handicaps_config_path, options.delimiter)

			rows = get_input_rows(options.exported_file_path, options.delimiter)
			
			if store_rows(rows, get_table_name(rows)) > 0:
			
				output_path = get_output_path(options)

				build_output(get_table_name(rows), output_path, options)
			
				open_output(output_path)
			else:
				logging.warning("there is no data to show")
			
			time.sleep(options.frequency)
		except (IOError, OSError), (errno, strerror):
			logging.error("error(%s): %s" % (errno, strerror))
			error = True
			time.sleep(options.frequency)

#-------------------------------    
if __name__ == "__main__":
	signal.signal(signal.SIGINT, signal_handler)
	main()
