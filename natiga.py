#!/usr/bin/python3
# -*- coding: utf-8 -*-

# Saves students (by seats and by school) results and ranks from http://natiga.nezakr.org

# Usage examples:
#   - ./natiga.py -f html --seats {20001..20023}
#   - ./natiga.py -f excel sqlite --school <link-to-school>

import bs4, argparse, sqlite3, json, sys, urllib3, openpyxl

http = None
students = []

class School:
    def __init__(self, link):
        try:
            self.link = link
            page = bs4.BeautifulSoup(open_link('POST', self.link,
                fields={'page': 1, 'order': 'mark_desc'}), 'lxml')
            res = page.find('tbody')
            if res == None: raise ValueError()
            # Page numbers
            num = page.find(attrs={'class': 'pagination'})
            #current = int(num.find(attrs={'class': 'active'}).text)
            last = int(num.findAll('li')[-1].text)
            # Registering students
            for page in range(2, last+1):
                students = res.findAll('tr')
                for student in students:
                    link = 'natiga.nezakr.org/%s' % student.findAll('td')[1].a['href']
                    Student(link=link)
                res = bs4.BeautifulSoup(open_link(self.link,
                    fields={'page': page, 'order': 'mark_desc'}), 'lxml').find('tbody')
        except ValueError:
            print('Invalid School Link')

class Student:
    def __init__(self, seat=None, link=None):
        try:
            if seat: link = 'natiga.nezakr.org/index.php?t=num&k=%d' % seat
            elif not link: raise AssertionError('A link or a seat number must be provided')
            page = bs4.BeautifulSoup(open_link('GET', link), 'lxml')
            res = page.findAll('tbody')
            if res == None: raise ValueError()
            # Student data
            self.info = {}
            data = res[0].findAll('td')
            i = 0
            while i < 11:
                self.info[data[i*2].text.strip()] = data[i*2+1].text.strip()
                if i in (2, 7): i += 3
                else: i += 1
            # Student marks
            self.marks = {}
            data = res[1].findAll('td')
            for i in range(0, 16):
                self.marks[data[i*3].text.strip()] = data[i*3+1].text.strip()
            # Student ranks
            data = res[2].findAll('td')
            self.ranks = {
                'الترتيب على الجمهورية': data[2].text.strip(),
                'الترتيب على الشعبة': data[5].text.strip(),
                'الترتيب على المحافظة': data[8].text.strip()
            }
            students.append(self)
        except ValueError:
            print('Invalid Seat Number: %d' % seat if seat else 'Invalid Student Link: %s' % link)

def parse_args():
    parser = argparse.ArgumentParser(description="Ranks students' results", epilog='(C) 2018 -- Amr Ayman')

    parser.add_argument('--seats', nargs='+', type=int, help='Student seat numbers')
    parser.add_argument('--schools', nargs='+', help='Link to a school')
    parser.add_argument('-o', '--outfile', required=True, help='Output filename')
    parser.add_argument('-f', default=['html'], nargs='+', choices=['html', 'excel', 'sqlite'],
        help='Output file format. You can specify multiple, e.g: -f html excel ..', dest='fileformats')
    options = parser.parse_args()
    # Options stuff
    if options.seats: options.seats = set(options.seats)
    elif options.schools: options.schools = set(options.schools)
    else: parser.error('No data given, add schools or seats')
    options.fileformats = set(options.fileformats)
    return options

def open_link(method, link, **kwargs):
    try: return http.request(method, link, redirect=False, **kwargs).data
    except Exception as e: print('Link cannot be open: %s' % e, file=sys.stderr)

if __name__ == '__main__':
    try:
        # Initializing environment
        http = urllib3.PoolManager(timeout=7, retries=4)
        options = parse_args()
        # Collecting data
        if options.schools:
            for school in options.schools:
                School(school)
        if options.seats:
            for seat in options.seats:
                Student(seat=seat)
        # Sorting according to marks. If one school only, sorted already.
        if options.seats or len(options.schools) > 1:
            students.sort(key=lambda student: float(student.info['المجموع']), reverse=True)
        # Writing data
        headers = []
        for h in students[0].info.keys(), students[0].marks.keys(), students[0].ranks.keys():
            headers.extend(h)
        for format in options.fileformats:
            if format == 'html':
                file = options.outfile + '.html'
                with open(file, 'w') as f:
                    f.write('<table><tr>')
                    for header in headers:
                        f.write('<th>%s</th>' % header)
                    f.write('</tr>')
                    for student in students:
                        f.write('<tr>')
                        for h in student.info.values(), student.marks.values(), student.ranks.values():
                            for v in h: f.write('<td>%s</td>' % v)
                        f.write('</tr>')
                    f.write('</table>')
            elif format == 'excel':
                file = options.outfile + '.xsl'
                wb = openpyxl.Workbook()
                wsl = wb.active
                for i, header in enumerate(headers, 1):
                    wsl.cell(row=1, column=i, value=header)
                for x, student in enumerate(students, 2):
                    data = []
                    for h in student.info.values(), student.marks.values(), student.ranks.values():
                        data.extend(h)
                    for i, datum in enumerate(data, 1):
                        wsl.cell(row=x, column=i, value=datum)
                wb.save(file)
            elif format == 'sqlite':
                file = options.outfile + '.db'
                conn = sqlite3.connect(file)
                c = conn.cursor()
                tmp = ['"%s" string' % header for header in headers]
                c.execute('create table results (%s)' % ', '.join(tmp))
                for student in students:
                    tmp = []
                    for h in student.info.values(), student.marks.values(), student.ranks.values():
                        tmp.extend(h)
                    c.execute('insert into results values (%s)' % ','.join('?'*len(tmp)), tmp)
                conn.commit()
                c.close()
                conn.close()
            print('Written to %s!' % file)
    except KeyboardInterrupt:
        print('\nInterrupted. Exiting ..', file=sys.stderr)
        sys.exit(1)
