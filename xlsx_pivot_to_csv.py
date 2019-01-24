import os
import zipfile
import logging
import pandas as pd
from shutil import copyfile
import xml.etree.ElementTree as et

filename = 'test/pivot_table_example.xlsx'
zipped_file = 'test/temp_zip_file.zip'

unzip_dir = 'unzip_tmp'
copyfile(filename, zipped_file)
zip_obj = zipfile.ZipFile(zipped_file, 'r')
zip_obj.extractall(unzip_dir)
zip_obj.close()

defdict = {}
columnas = []
definitions = unzip_dir + '/xl/pivotCache/pivotCacheDefinition1.xml'
e = et.parse(definitions).getroot()
for fields in e.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cacheFields'):
    for cidx, field in enumerate(list(fields)):
        columna = field.attrib.get('name')
        defdict[cidx] = []
        columnas.append(columna)
        for value in list(list(field)[0]):
            tagname = value.tag
            defdict[cidx].append(value.attrib.get('v', 0))

dfdata = []
bdata = unzip_dir + '/xl/pivotCache/pivotCacheRecords1.xml'
estr = 'this it not should happen index cidx = {} vattrib = {} defaultidcts = {} tmpdata for the time = {} xml raw {}'
for event, elem in et.iterparse(bdata, events=('start', 'end')):
    if elem.tag == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}r' and event == 'start':
        tmpdata = []
        for cidx, valueobj in enumerate(list(elem)):
            tagname = valueobj.tag
            vattrib = valueobj.attrib.get('v')
            rdata = vattrib
            if tagname == '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}x':
                try:
                    rdata = defdict[cidx][int(vattrib)]
                except Exception:
                    logging.error(
                        estr.format(cidx, vattrib, defdict, tmpdata, et.tostring(elem, encoding='utf8', method='xml')))
            tmpdata.append(rdata)
        if tmpdata:
            dfdata.append(tmpdata)
        elem.clear()

df = pd.DataFrame(dfdata)
new_filename = os.path.splitext(filename)[0] + '.csv'
df.to_csv(new_filename, index=False)
