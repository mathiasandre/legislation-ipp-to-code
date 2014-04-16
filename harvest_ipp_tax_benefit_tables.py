#! /usr/bin/env python
# -*- coding: utf-8 -*-


# Law-to-Code -- Extract formulas & parameters from laws
# By: Emmanuel Raviart <emmanuel@raviart.com>
#
# Copyright (C) 2013, 2014 OpenFisca Team
# https://github.com/openfisca/LawToCode
#
# This file is part of Law-to-Code.
#
# Law-to-Code is free software; you can redistribute it and/or modify
# it under the terms of the GNU Affero General Public License as
# published by the Free Software Foundation, either version 3 of the
# License, or (at your option) any later version.
#
# Law-to-Code is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU Affero General Public License for more details.
#
# You should have received a copy of the GNU Affero General Public License
# along with this program.  If not, see <http://www.gnu.org/licenses/>.


"""Extract parameters from IPP's tax benefit tables.

Note: Currently this script requires an XLS version of the tables. XLSX file must be converted to XLS before use.

IPP = Institut des politiques publiques
http://www.ipp.eu/en/tools/ipp-tax-and-benefit-tables/
http://www.ipp.eu/fr/outils/baremes-ipp/
"""


import argparse
import collections
import itertools
import json
import logging
import os
import sys
import urlparse

from biryani1 import baseconv, custom_conv, datetimeconv, states
import xlrd

app_name = os.path.splitext(os.path.basename(__file__))[0]
conv = custom_conv(baseconv, datetimeconv, states)
date_color_index = 41
heading_color_index = 31
log = logging.getLogger(app_name)
N_ = lambda message: message
note_color_index = 31
parameters = []


currency_converter = conv.first_match(
    conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        conv.test_none(),
        ),
    conv.pipe(
        conv.test_isinstance(tuple),
        conv.test(lambda couple: len(couple) == 2, error = N_(u"Invalid couple length")),
        conv.struct(
            (
                conv.pipe(
                    conv.test_isinstance((float, int)),
                    conv.not_none,
                    ),
                conv.pipe(
                    conv.test_isinstance(basestring),
                    conv.test_in([
                        u'EUR',
                        u'FRF',
                        ]),
                    ),
                ),
            ),
        ),
    )


pss_converters = collections.OrderedDict((
    (u"Date d'effet", conv.pipe(
        conv.test_isinstance(basestring),
        conv.iso8601_input_to_date,
        conv.date_to_iso8601_str,
        conv.not_none,
        )),
    (u'Plafond de la Sécurité sociale (mensuel)', currency_converter),
    (u'Plafond de la Sécurité sociale (annuel)', currency_converter),
    (u'Référence législative', conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        )),
    (u'Parution au JO', conv.pipe(
        conv.test_isinstance(basestring),
        conv.iso8601_input_to_date,
        conv.date_to_iso8601_str,
        )),
    (u'Notes', conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        )),
    (None, conv.pipe(
        conv.test_isinstance(basestring),
        conv.cleanup_line,
        conv.test_none(),
        )),
    ))


def enumerate_row_cells_type_and_value(sheet, row_index, cell_coordinates_by_merged_coordinates):
    cell_coordinates_by_merged_column_index = cell_coordinates_by_merged_coordinates.get(row_index, {})
    for column_index in range(len(sheet.row_types(row_index))):
        unmerged_cell_coordinates = cell_coordinates_by_merged_column_index.get(column_index)
        if unmerged_cell_coordinates is None:
            unmerged_row_index = row_index
            unmerged_column_index = column_index
        else:
            unmerged_row_index, unmerged_column_index = unmerged_cell_coordinates
        yield (
            column_index,
            sheet.row_types(unmerged_row_index)[unmerged_column_index],
            sheet.row_values(unmerged_row_index)[unmerged_column_index],
            )


def get_unmerged_cell_coordinates(row_index, column_index, cell_coordinates_by_merged_coordinates):
    unmerged_cell_coordinates = cell_coordinates_by_merged_coordinates.get(row_index, {}).get(column_index)
    if unmerged_cell_coordinates is None:
        return row_index, column_index
    return unmerged_cell_coordinates


def main():
    parser = argparse.ArgumentParser(description = __doc__)
    parser.add_argument('xls_path', help = 'path of XLS file')
    parser.add_argument('-v', '--verbose', action = 'store_true', default = False, help = "increase output verbosity")
    args = parser.parse_args()
    logging.basicConfig(level = logging.DEBUG if args.verbose else logging.WARNING, stream = sys.stdout)

    book = xlrd.open_workbook(filename = args.xls_path, formatting_info = True)
    sheet_names = book.sheet_names()
    sheet_names = [
        sheet_name
        for sheet_name in book.sheet_names()
        if not sheet_name.startswith((u'Sommaire', u'Outline', u'Barème IGR'))
        ]

    for sheet_name in sheet_names:
        log.info('Parsing sheet {}'.format(sheet_name)
        sheet = book.sheet_by_name(sheet_name)

        # Extract coordinates of merged cells.
        cell_coordinates_by_merged_coordinates = {}
        for row_low, row_high, column_low, column_high in sheet.merged_cells:
            for row_index in range(row_low, row_high):
                cell_coordinates_by_merged_column_index = cell_coordinates_by_merged_coordinates.setdefault(
                    row_index, {})
                for column_index in range(column_low, column_high):
                    cell_coordinates_by_merged_column_index[column_index] = (row_low, column_low)

        descriptions_rows = []
        labels_rows = []
        notes_row = None
        state = None
        taxipp_names_row = None
        values_rows = []
        for row_index in range(sheet.nrows):
            if state is None:
                taxipp_names_row = [
                    transform_xls_cell_to_str(book, cell_type, cell_value, sheet.cell_xf_index(row_index, column_index))
                    for column_index, cell_type, cell_value in enumerate_row_cells_type_and_value(sheet, row_index,
                        cell_coordinates_by_merged_coordinates)
                    ]
                state = 'labels'
                continue
            if state == 'labels':
                xf_index = sheet.cell_xf_index(row_index, 0)
                xf = book.xf_list[xf_index]  # gets an XF object
                if xf.background.background_colour_index == heading_color_index:
                    labels_rows.append([
                        transform_xls_cell_to_str(book, cell_type, cell_value,
                            sheet.cell_xf_index(row_index, column_index))
                        for column_index, cell_type, cell_value in enumerate_row_cells_type_and_value(sheet, row_index,
                            cell_coordinates_by_merged_coordinates)
                        ])
                    continue
                state = 'values'
            if state == 'values':
                xf_index = sheet.cell_xf_index(*get_unmerged_cell_coordinates(row_index, 0,
                    cell_coordinates_by_merged_coordinates))
                xf = book.xf_list[xf_index]  # gets an XF object
                if xf.background.background_colour_index == date_color_index:
                    values_rows.append([
                        transform_xls_cell_to_json(book, cell_type, cell_value,
                            sheet.cell_xf_index(row_index, column_index))
                        for column_index, cell_type, cell_value in enumerate_row_cells_type_and_value(sheet, row_index,
                            cell_coordinates_by_merged_coordinates)
                        ])
                    continue
                state = 'notes'
            if state == 'notes':
                xf_index = sheet.cell_xf_index(*get_unmerged_cell_coordinates(row_index, 0,
                    cell_coordinates_by_merged_coordinates))
                xf = book.xf_list[xf_index]  # gets an XF object
                if xf.background.background_colour_index == note_color_index:
                    assert notes_row is None
                    notes_row = [
                        transform_xls_cell_to_str(book, cell_type, cell_value,
                            sheet.cell_xf_index(row_index, column_index))
                        for column_index, cell_type, cell_value in enumerate_row_cells_type_and_value(sheet, row_index,
                            cell_coordinates_by_merged_coordinates)
                        ]
                    continue
                state = 'description'
            assert state == 'description'
            descriptions_rows.append([
                transform_xls_cell_to_str(book, cell_type, cell_value,
                    sheet.cell_xf_index(row_index, column_index))
                for column_index, cell_type, cell_value in enumerate_row_cells_type_and_value(sheet, row_index,
                    cell_coordinates_by_merged_coordinates)
                ])
        description = u'\n'.join(
            u' '.join(
                cell.strip()
                for cell in row
                if cell is not None
                )
            for row in descriptions_rows
            ) or None
        print taxipp_names_row
        print labels_rows
        print values_rows
        print notes_row
        print description
        boum

#        for row in sheet.rows:
#            if state is None:
#                taxipp_names = [
#                    cell.value
#                    for cell in row
#                    ]
#                for index, value in reversed(list(enumerate(taxipp_names))):
#                    if value is not None:
#                        del taxipp_names[index + 1:]
#                        break
#                state = 'labels'
#                continue
#            if state == 'labels':
#                print row[0].style
#                print row[0].style.fill.fill_type
#                print row[0].style.fill.start_color
#                print row[0].style.fill.end_color
#                labels = [
#                    cell.value
#                    for cell in row
#                    ]
#                for index, value in reversed(list(enumerate(labels))):
#                    if value is not None:
#                        del labels[index + 1:]
#                        break
#                break
#        print taxipp_names
#        print labels
#        boum
#        taxipp_names = sheet.rows

    sheet = book.sheet_by_name(u'PSS')
    sheet_data = [
        [
            transform_xls_cell_to_json(book, cell_type, cell_value, sheet.cell_xf_index(row_index, column_index))
            for column_index, (cell_type, cell_value) in enumerate(itertools.izip(sheet.row_types(row_index),
                sheet.row_values(row_index)))
            ]
        for row_index in range(sheet.nrows)
        ]
    taxipp_names = sheet_data[0]
    labels = sheet_data[1]
    assert labels == pss_converters.keys(), str((labels,))
    taxipp_name_by_label = dict(zip(labels, taxipp_names))
    description_lines = []
    entries = []
    state = None
    for row_index, row in enumerate(itertools.islice(sheet_data, 2, None)):
        if all(cell in (None, u'') for cell in row):
            state = 'description'
        if state is None:
            entry = conv.check(conv.struct(pss_converters))(dict(zip(labels, row)), state = conv.default_state)
            entries.append(entry)
        else:
            description_line = u' '.join(
                cell.strip()
                for cell in row
                if cell is not None
                )
            description_lines.append(description_line)
    description = u'\n'.join(description_lines) or None

    parameters = []
    for entry in entries:
        value_label = u'Plafond de la Sécurité sociale (mensuel)'
        parameters.append(dict(
            comment = entry[u"Notes"],
            description = description,
            format = u'float',
            legislative_reference = entry[u'Référence législative'],
            official_publication_date = entry[u'Parution au JO'],
            start_date = entry[u"Date d'effet"],
            taxipp_code = taxipp_name_by_label[value_label],
            title = value_label,
            unit = entry[value_label][1]
                if entry[value_label] is not None
                else None,
            value = entry[value_label][0]
                if entry[value_label] is not None
                else None,
            ))
        value_label = u'Plafond de la Sécurité sociale (annuel)'
        parameters.append(dict(
            comment = entry[u"Notes"],
            description = description,
            format = u'float',
            legislative_reference = entry[u'Référence législative'],
            official_publication_date = entry[u'Parution au JO'],
            start_date = entry[u"Date d'effet"],
            taxipp_code = taxipp_name_by_label[value_label],
            title = value_label,
            unit = entry[value_label][1] if entry[value_label] is not None else None,
            value = entry[value_label][0] if entry[value_label] is not None else None,
            ))

    parameter_upsert_url = urlparse.urljoin(conf['law_to_code.site_url'], 'api/1/parameters/upsert')
    for parameter in parameters:
        response = requests.post(parameter_upsert_url,
            data = unicode(json.dumps(dict(
                api_key = conf['law_to_code.api_key'],
                value = parameter,
                ), ensure_ascii = False, indent = 2)).encode('utf-8'),
            headers = {
                'Content-Type': 'application/json; charset=utf-8',
                'User-Agent': conf['user_agent']
                }
            )
        if not response.ok:
            print response.json()
            response.raise_for_status()

    return 0


def transform_xls_cell_to_json(book, type, value, xf_index):
    """Convert an XLS cell (type & value) to an unicode string.

    Code taken from http://code.activestate.com/recipes/546518-simple-conversion-of-excel-files-into-csv-and-yaml/

    Type Codes:
    EMPTY   0
    TEXT    1 a Unicode string
    NUMBER  2 float
    DATE    3 float
    BOOLEAN 4 int; 1 means TRUE, 0 means FALSE
    ERROR   5
    """
    if type == 0:
        value = None
    elif type == 1:
        if not value:
            value = None
    elif type == 2:
        # NUMBER
        value_int = int(value)
        if value_int == value:
            value = value_int
        xf = book.xf_list[xf_index] # gets an XF object
        format_key = xf.format_key
        format = book.format_map[format_key] # gets a Format object
        format_str = format.format_str # this is the "number format string"
        if format_str == 'GENERAL':
            return value
        if format_str.endswith(ur'\ "€"'):
            return (value, u'EUR')
        if format_str.endswith(ur'\ [$FRF]'):
            return (value, u'FRF')
        if format_str.endswith(u'%'):
            return (value, u'%')
        print value, format_str
        TODO
    elif type == 3:
        # DATE
        y, m, d, hh, mm, ss = xlrd.xldate_as_tuple(value, book.datemode)
        date = u'{0:04d}-{1:02d}-{2:02d}'.format(y, m, d) if any(n != 0 for n in (y, m, d)) else None
        value = u'T'.join(
            fragment
            for fragment in (
                date,
                u'{0:02d}:{1:02d}:{2:02d}'.format(hh, mm, ss)
                    if any(n != 0 for n in (hh, mm, ss)) or date is None
                    else None,
                )
            if fragment is not None
            )
    elif type == 4:
        value = bool(value)
    elif type == 5:
        # ERROR
        value = xlrd.error_text_from_code[value]
    return value


def transform_xls_cell_to_str(book, type, value, xf_index):
    cell = transform_xls_cell_to_json(book, type, value, xf_index)
    assert cell is None or isinstance(cell, basestring), u'Expected a string. Got: {}'.format(cell).encode('utf-8')
    return cell


if __name__ == "__main__":
    sys.exit(main())
