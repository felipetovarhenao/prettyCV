import os
import json
import subprocess
from docx import Document
from time import sleep
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from docx.shared import Pt, Inches
from datetime import date
from collections.abc import Callable
from utils import (
    insertHR,
    format_date_range,
    add_hyperlink,
    format_year_range,
    parse_date,
    indent_table,
    add_page_number,
    BLUE,
    DARK_BLUE,
    GREY,
    BLACK
)


class CV:
    """ Curriculum Vitae class """
    CV_KEY = 'cv'
    WORKS_KEY = 'works'

    def __init__(self, font: str = 'Nunito', reverse_format: bool = False) -> None:
        self.doc = Document()
        self.font = font
        self.font_size = 10.5
        self.tab_size = Inches(0.25)
        self.date_col_width = Inches(0.95)
        self.item_col_width = Inches(5.25)
        self.reverse_format = reverse_format
        for style_name in ['Normal'] + [f'Heading {i}' for i in range(1, 6)]:
            style = self.doc.styles[style_name]
            style.font.name = self.font
            style.font.size = Pt(self.font_size)
            style.font.color.rgb = BLACK
            style.paragraph_format.space_after = Pt(0)
            style.paragraph_format.space_before = Pt(0)
        add_page_number(self.doc.sections[0].footer.paragraphs[0].add_run())
        self.doc.sections[0].footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    def write(self, output_file: str) -> None:
        self.doc.save(output_file)
        cmd = """osascript -e 'tell application "Microsoft Word" to close windows'"""
        os.system(cmd)
        sleep(1)
        subprocess.run(['open', output_file])

    def load_data(self, cv_path: str, works_path: str) -> None:
        data = {}
        for json_file, key in zip([cv_path, works_path], [self.CV_KEY, self.WORKS_KEY]):
            with open(json_file, 'r') as f:
                data[key] = json.load(f)
        self.data = data

    def __new_section(self, name: str) -> Paragraph:
        self.__insert_break(2)
        header = self.doc.add_heading(name)
        header.runs[0].font.name = self.font
        insertHR(header)
        return header

    def __new_subsection(self, name: str) -> Paragraph:
        self.__insert_break()
        subheader = self.doc.add_heading(name, level=2)
        font = subheader.runs[0].font
        font.name = self.font
        font.color.rgb = DARK_BLUE
        self.__insert_break()
        return subheader

    def __insert_break(self, n_units: int | float = 1, parent=None):
        obj = parent or self.doc
        gap = Pt(7 * n_units)
        p = obj.add_paragraph(" ")
        p.runs[0].font.size = gap
        p.paragraph_format.line_spacing = gap

    def compile(self):
        self.parse_basics()
        self.parse_education()
        self.parse_work_experience()
        self.parse_publications()
        self.parse_awards()
        self.parse_skills()
        self.parse_works()

    def parse_basics(self):
        basics = self.data[self.CV_KEY]['basics']

        p = self.doc.add_paragraph("")
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER

        name = p.add_run(basics['name'].upper())
        name.bold = True

        labels = p.add_run("\n{} | Curriculum Vitae".format(" + ".join([l.capitalize() for l in basics['labels']])))
        labels.bold = True

        address_key = p.add_run("\nAddress: ")
        address_key.bold = True

        p.add_run("{}, {}, {} — {}. ".format(*[basics['location'][x]
                                               for x in ['address', 'city', 'region', 'countryCode',]]))

        phone_key = p.add_run("Phone: ")
        phone_key.bold = True

        _phone = basics['phone']
        p.add_run(_phone)

        _url = basics['profiles'][-1]['url']
        p.add_run(f"\n")
        add_hyperlink(p, _url, _url)
        p.add_run(f" | ")
        add_hyperlink(p, basics['email'], basics['email'])

    def __make_entry_table(self, parent: object, items: list, handler: Callable, date_getter: Callable) -> None:
        tbl = parent.add_table(len(items), 2)
        last_date = None
        for i, item in enumerate(items):
            date_cell, item_cell = tbl.cell(i, int(self.reverse_format)), tbl.cell(i, 1-int(self.reverse_format))
            date_cell.width, item_cell.width = self.date_col_width, self.item_col_width
            date = date_getter(item)
            if date != last_date:
                date_cell.paragraphs[0].add_run(date)
            last_date = date
            last_element = handler(cell=item_cell, item=item)
            self.__insert_break(0.5, parent=last_element or date_cell)

    def parse_education(self):
        education = self.data[self.CV_KEY].pop('education', None)
        if not education:
            return
        self.__new_section("EDUCATION")
        degrees = education.pop('degrees', None)
        if degrees:
            self.__new_subsection("Degrees")
            for degree in degrees:
                p = self.doc.add_paragraph()
                name = p.add_run(degree['name'])
                name.bold = True

                p.add_run(f", {degree['major']}.")

                p = self.doc.add_paragraph()
                p.paragraph_format.left_indent = self.tab_size

                institution = p.add_run(f"{degree['institution']}")
                institution.bold = True

                p.add_run(f". {degree['city']}. {degree['country']}. {format_year_range(*degree['date'])}.")
                minors = degree['minors']
                if minors:
                    minors_label = p.add_run("\nMinor fields: ")
                    minors_label.font.bold = True
                    p.add_run(f"{', '.join(degree['minors'])}.")

        other_ed = education.pop('other', None)
        if other_ed:
            self.__new_subsection("Other")
            for ed in other_ed:
                p = self.doc.add_paragraph("")
                p.paragraph_format.left_indent = self.tab_size
                p.paragraph_format.first_line_indent = -self.tab_size

                url = ed['url']
                if url:
                    name = add_hyperlink(p, ed['name'], url)
                else:
                    name = p.add_run({ed['name']})
                name.bold = True
                p.add_run(f" ({ed['type']}). {ed['institution']}. {ed['location']}. {format_date_range(*ed['date'])}.")

    def parse_work_experience(self):
        self.parse_jobs()
        self.parse_lectures()
        self.parse_workshops()
        self.parse_residencies()

    def parse_jobs(self):
        def handler(cell, item):
            position = item
            p = cell.paragraphs[0]
            name = p.add_run(position['name'])
            name.bold = True

            p.add_run(f". {position['workplace']}. {position['city']}. {position['country']}.")

            _courses = position.pop('courses', None)
            if _courses:
                num_courses = len(_courses)
                course_tbl = cell.add_table(num_courses, 2)
                course_label_cell = course_tbl.cell(0, 0)
                course_label_cell.paragraphs[0].add_run("Courses:")

                for j, course in enumerate(_courses):
                    course_cell = course_tbl.cell(j, 1)
                    course_cell.width = Inches(20)
                    cp = course_cell.paragraphs[0]
                    course_name = cp.add_run(course['name'])
                    course_name.font.bold = True
                    terms = cp.add_run(", {}.".format(course['terms']))
                    terms.font.color.rgb = GREY
                return course_cell

        def date_getter(item):
            return format_year_range(*item['date'])

        work = self.data[self.CV_KEY]['work']
        self.__new_section("WORK EXPERIENCE")
        academic_work = work.pop('academic', None)
        other_work = work.pop('other positions', None)
        experience_sections = [x for x in [("Teaching experience", academic_work),
                                           ("Other positions", other_work)] if x[1] is not None]
        for label, positions in experience_sections:
            self.__new_subsection(label)
            positions.sort(key=lambda x: 10000 if x['date'][1] ==
                           True else (x['date'][0] if x['date'][1] == False else x['date'][1]), reverse=True)
            self.__make_entry_table(self.doc, positions, handler, date_getter)

    def parse_lectures(self):
        lectures = self.data[self.CV_KEY]['work'].pop('lectures', None)
        if not lectures:
            return
        self.__new_subsection("Guest lectures")
        lectures.sort(key=lambda x: sorted(x['events'], key=lambda y: y['date'], reverse=True)[0]['date'], reverse=True)
        for lecture in lectures:
            p = self.doc.add_paragraph()
            name = p.add_run(lecture['name'])
            name.bold = True
            name.italic = True

            for event in lecture['events']:
                p = self.doc.add_paragraph("@ ")
                p.runs[0].font.color.rgb = BLUE
                p.paragraph_format.left_indent = self.tab_size
                event_name = p.add_run(event['name'])
                event_name.font.italic = True
                year, month, day = parse_date(event['date'])
                p.add_run(f". { event['venue']}. {event['city']}. {event['country']}. {f'{month} {day}, {year}'}.")

    def parse_workshops(self):
        workshops = self.data[self.CV_KEY]['work'].pop('workshops', None)
        if not workshops:
            return
        self.__new_subsection("Workshops")
        workshops.sort(key=lambda x: sorted(x['events'], key=lambda y: y['date'], reverse=True)[0]['date'], reverse=True)

        for workshop in workshops:
            p = self.doc.add_paragraph()
            name = p.add_run(workshop['name'])
            name.bold = True
            name.italic = True

            for event in workshop['events']:
                p = self.doc.add_paragraph("@ ")
                p.runs[0].font.color.rgb = BLUE
                p.paragraph_format.left_indent = self.tab_size
                institution = p.add_run(event['institution'])
                institution.italic = True
                year, month, day = parse_date(event['date'])

                p.add_run(
                    f". {event['numSessions']} sessions ({event['totalHours']} hours total). {event['city']}. {event['country']}. {month} {day}, {year}.")

    def parse_residencies(self):
        residencies = self.data[self.CV_KEY]['work'].pop('residencies', None)
        if not residencies:
            return

        def date_getter(item):
            return parse_date(item['date'])[0]

        def handler(cell, item):
            residency = item
            p = cell.paragraphs[0]
            role = p.add_run(residency['role'])
            role.font.italic = True
            at = p.add_run(" @ ")
            at.font.color.rgb = BLUE
            event = p.add_run(residency['event'])
            event.font.bold = True

            _institution = residency['institution']
            _end = residency['end']
            date_range = format_date_range(residency['date'], _end)
            p.add_run(f". {_institution}. {date_range}.")

            p = cell.add_paragraph()
            p.paragraph_format.left_indent = self.tab_size
            label = p.add_run("Activities: ")
            label.font.bold = True
            _activities = "{}.".format(", ".join(residency['activities']))
            activities = p.add_run(_activities)
            activities.font.italic = True

        self.__new_subsection("Residencies")
        residencies.sort(key=lambda x: x['date'], reverse=True)
        self.__make_entry_table(self.doc, residencies, handler, date_getter)

    def parse_awards(self):
        self.__new_section("AWARDS")
        awards = []
        commissions = []
        for work in self.data[self.WORKS_KEY]:
            work_awards = work['awards']
            commission = work['commission']
            if work_awards:
                for award in work_awards:
                    awards.append(award)
            if commission:
                commissions.append(work)

        self._parse_awards(awards, "Artistic awards")
        self._parse_commissions(commissions)
        self._parse_awards(self.data[self.CV_KEY]['awards'].pop('academic', None), "Academic awards")

    def _parse_commissions(self, commissions):
        if not commissions:
            return

        def handler(cell, item):
            commission = item
            p = cell.paragraphs[0]
            name = p.add_run(commission['name'])
            name.font.bold = True

            subtitle = p.add_run(" {}.".format(commission['subtitle']))
            subtitle.font.italic = True

            p.add_run(" {}.".format(commission['commission']))

        def date_getter(item):
            return str(item['year'])
        self.__new_subsection("Commissions")
        commissions.sort(key=lambda x: x['year'], reverse=True)
        self.__make_entry_table(self.doc, commissions, handler, date_getter)

    def _parse_awards(self, awards, label):
        if not awards:
            return

        def handler(cell, item):
            award = item
            p = cell.paragraphs[0]
            name = p.add_run(award['name'])
            name.font.bold = True

            p.add_run(". {}. {}.".format(*[award[x] for x in ['institution', 'country']]))

        def date_getter(item):
            return str(item['date'])

        self.__new_subsection(label)
        awards.sort(key=lambda x: x['date'], reverse=True)
        self.__make_entry_table(self.doc, awards, handler, date_getter)

    def parse_publications(self):
        publications = self.data[self.CV_KEY]['work'].pop('publications', None)
        if publications:
            def handler(cell, item):
                pub = item
                p = cell.paragraphs[0]
                p.add_run(f"{pub['author']} ({pub['date']}). ")
                name = p.add_run(pub['name'])
                name.font.bold = True

                publisher = p.add_run(f". {pub['publisher']}")
                publisher.font.italic = True

                p.add_run(f", ({pub['edition']}), {'-'.join([str(x) for x in pub['pages']])}. ")
                add_hyperlink(p, pub['doi'], pub['doi'])

            def date_getter(item):
                return str(item['date'])

            self.__new_section("PUBLICATIONS")
            articles, scores = publications[:1], publications[1:]
            for items, label in [(articles, 'Peer-reviewed articles'), (scores, 'Scores')]:
                self.__new_subsection(label)
                items.sort(key=lambda x: x['date'])
                self.__make_entry_table(self.doc, items, handler, date_getter)

        software_list = self.data[self.CV_KEY]['work'].pop('software', None)
        if software_list:
            def handler(cell, item):
                software = item
                p = cell.paragraphs[0]
                name = p.add_run(software['name'])
                name.font.bold = True

                url = software['url']
                p.add_run(" (")
                add_hyperlink(p, url, url)
                p.add_run(")")

                p = cell.add_paragraph("Keywords: ")
                p.runs[0].italic = p.runs[0].bold = True
                p.paragraph_format.left_indent = self.tab_size
                keywords = p.add_run(f'{", ".join(software["keywords"])}.')
                keywords.italic = True

                p = cell.add_paragraph("Description: ")
                p.runs[0].italic = p.runs[0].bold = True
                p.paragraph_format.left_indent = self.tab_size
                descr = p.add_run(f"{software['description']}")
                descr.font.italic = True

            def date_getter(item):
                return str(item['year'])

            self.__new_subsection("Software contributions")
            software_list.sort(key=lambda x: x['year'], reverse=True)
            self.__make_entry_table(self.doc, software_list, handler, date_getter)

    def parse_skills(self):
        self.__new_section("SKILLS")
        skills_dict = self.data[self.CV_KEY].pop('skills', None)
        if not skills_dict:
            return
        divs = 3
        for skill_key in skills_dict:
            self.__new_subsection(skill_key.capitalize())
            skills = skills_dict[skill_key]
            skills.sort(key=lambda x: x['level'])
            tbl = self.doc.add_table(round(len(skills) / divs), divs)
            indent_table(tbl, 350)
            for i, skill in enumerate(skills):
                row = i // divs
                col = i % divs
                cell = tbl.cell(row, col)

                p = cell.paragraphs[0]

                name = p.add_run(skill['name'])
                name.font.bold = True

                keywords = p.add_run(f' ({", ".join(skill["keywords"])})')
                keywords.font.italic = True

                gap = Pt(5)

                p = cell.add_paragraph(" ")
                p.runs[0].font.size = gap
                p.paragraph_format.line_spacing = gap
            self.__insert_break()

    def parse_works(self):
        works = self.data[self.WORKS_KEY]
        if not works:
            return
        self.doc.add_page_break()
        self.__new_section("LIST OF WORKS")
        works.sort(key=lambda x: x['year'], reverse=True)
        last_date = None
        for i, work in enumerate(works):
            date = str(work['year'])
            if date != last_date:
                self.__new_subsection(f"{date}")
            last_date = date
            p = self.doc.add_paragraph()
            name = p.add_run(work['name'])
            name.font.bold = True

            p.add_run(f" ({date}) ")
            subtitle = p.add_run(f"{work['subtitle']}. ")
            subtitle.italic = True

            p.add_run(f"{work['duration']}'")

            _commission = work['commission']
            if _commission:
                self.__insert_break(0.5)
                p = self.doc.add_paragraph()
                p.paragraph_format.left_indent = self.tab_size
                commission = p.add_run(_commission)
                commission.italic = True
                commission.font.color.rgb = GREY

            performances = work['performances']
            if performances:
                performances.sort(key=lambda x: x['date'], reverse=True)
                num_perf = len(performances)
                self.__insert_break(0.5)
                p = self.doc.add_paragraph("Performances")
                label = p.runs[0]
                label.italic = True
                label.bold = True
                label.font.color.rgb = DARK_BLUE
                p.paragraph_format.left_indent = self.tab_size
                self.__insert_break(0.5)
                for i, performance in enumerate(performances):
                    p = self.doc.add_paragraph()
                    p.paragraph_format.left_indent = self.tab_size * 2

                    event = p.add_run(performance['event'])
                    event.font.bold = True

                    if i == num_perf - 1:
                        p.add_run(" (world premiere)")

                    at = p.add_run(" @ ")
                    at.font.color.rgb = BLUE

                    year, month, day = parse_date(performance['date'])
                    p.add_run(
                        f"{performance['venue']}. {performance['city']}. {performance['country']}. {month} {day}, {year}. ")

                    performers = performance.pop('performers', None)
                    if performers:
                        p = self.doc.add_paragraph("Performed by ")
                        p.paragraph_format.left_indent = self.tab_size * 3
                        num_performers = len(performers)
                        for i, performer in enumerate(performers):
                            p.add_run(performer['name'])
                            role = p.add_run(
                                f" ({performer['role']}){'.' if i == num_performers - 1 else (', and ' if i == num_performers - 2 else ', ')}")
                            role.italic = True
                    self.__insert_break(0.5)


cv = CV()
cv_path = '../personal-website/src/json/cv.json'
works_path = '../personal-website/src/json/work-catalog.json'
cv.load_data(cv_path=cv_path, works_path=works_path)
cv.compile()
file_id = date.today()
cv.write(f'CV_{file_id}.docx')
