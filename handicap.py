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

# constants
DB_PATH = "db/handicap.db"
FONTS_PATH = "fonts"
DEF_OUTPUT_PATH = "output/handicap.pdf"
DEF_DELIMITER = ";"
DEF_CHECK_FREQUENCY = 5

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
	pdfmetrics.registerFont(TTFont('DejaVuSerif', '%s/DejaVuSerif.ttf' % FONTS_PATH))
	pdfmetrics.registerFont(TTFont('DejaVuSerif-Bold', '%s/DejaVuSerif-Bold.ttf' % FONTS_PATH))

	styles = getSampleStyleSheet()
	styles.add(ParagraphStyle(name='HNormal',
                               fontName='DejaVuSerif',
                               fontSize=12))
	
	styles.add(ParagraphStyle(name='HBold',
                               fontName='DejaVuSerif-Bold',
                               fontSize=12))
	return styles

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

def store_rows(rows, table_name):
	conn = get_connection()
	conn.text_factory = lambda x: unicode(x, 'utf-8', 'ignore')
	c = conn.cursor()

	c.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='%s'" % table_name)
	if len(c.fetchall()) != 1:
		c.execute("CREATE TABLE %s(%s)" % (table_name, (','.join(["%s text" % s for s in rows[0]])).replace('.', '_').replace(' ', '_')))
	else:
		c.execute("DELETE FROM %s" % table_name)

	rows.pop(0)
	
	c.executemany("INSERT INTO %s VALUES (%s)" % (table_name, ','.join(["?" for s in rows[0]])), rows)
	
	conn.commit()

	conn.close()

def get_rows(table_name):
	conn = get_connection()
	c = conn.cursor()
	c.execute('SELECT * FROM %s' % table_name)
	rows = c.fetchall()
	conn.close()
	return rows

def main(m_args=None):
	(options, args) = parse_arguments(m_args)

	styles = init_styles()
	
	moddate_old = 0
	error = False

	while True:
		try:
			
			moddate_new = os.stat(args[0])[8]
			if moddate_new == moddate_old and not error:
				time.sleep(options.frequency)
				continue
			else:
				moddate_old = moddate_new
				error = False
			
			doc = SimpleDocTemplate(options.output_path,pagesize=landscape(A4),
                        rightMargin=72,leftMargin=72,
                        topMargin=72,bottomMargin=18)

			story = []
			rows = get_input_rows(args[0], options.delimiter)
			store_rows(rows, 'rows')
			
			story.append(Paragraph(u'Nadpisčťž', styles["HBold"]))
			story.append(Paragraph('Nadpis 1', styles["HNormal"]))
			t = Table(get_rows('rows'))
			t.setStyle(TableStyle([
				('FONTNAME', (0, 0), (-1, -1), 'DejaVuSerif'),
				('FONTSIZE', (0, 0), (-1, -1), 12),
				('ALIGN',(0,0),(-1,-1),'LEFT')
				]))
			story.append(t)
			doc.build(story)
			
			webbrowser.open(r"file:///" + os.getcwd() + "/" + options.output_path)

			time.sleep(options.frequency)
		except (IOError, OSError), (errno, strerror):
			print "Error(%s): %s" % (errno, strerror)
			error = True
			time.sleep(options.frequency)

#-------------------------------    
if __name__ == "__main__":
	signal.signal(signal.SIGINT, signal_handler)
	main()
