from copy import deepcopy
import warnings
from lxml.etree import Element
from lxml import etree
from zipfile import ZipFile, ZIP_DEFLATED
import shlex
import re


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
)

CONTENT_TYPE_SETTINGS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'


class MailMerge(object):
    def __init__(self, file, remove_empty_tables=False):
        self.zip = ZipFile(file)
        self.parts = {}
        self.settings = None
        self._settings_info = None
        self.remove_empty_tables = remove_empty_tables

        try:
            content_types = etree.parse(self.zip.open('[Content_Types].xml'))
            for file in content_types.findall('{%(ct)s}Override' % NAMESPACES):
                type = file.attrib['ContentType' % NAMESPACES]
                if type in CONTENT_TYPES_PARTS:
                    zi, self.parts[zi] = self.__get_tree_of_file(file)
                elif type == CONTENT_TYPE_SETTINGS:
                    self._settings_info, self.settings = self.__get_tree_of_file(file)

            to_delete = []

            for part in self.parts.values():

                for parent in part.findall('.//{%(w)s}fldSimple/..' % NAMESPACES):
                    for idx, child in enumerate(parent):
                        if child.tag != '{%(w)s}fldSimple' % NAMESPACES:
                            continue
                        instr = child.attrib['{%(w)s}instr' % NAMESPACES]

                        name = self.__parse_instr(instr)
                        if name is None:
                            continue
                        parent[idx] = Element('MergeField', name=name)

                for parent in part.findall('.//{%(w)s}instrText/../..' % NAMESPACES):
                    children = list(parent)
                    fields = zip(
                        [children.index(e) for e in
                         parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]/..' % NAMESPACES)],
                        [children.index(e) for e in
                         parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="end"]/..' % NAMESPACES)]
                    )

                    for idx_begin, idx_end in fields:
                        # consolidate all instrText nodes between'begin' and 'end' into a single node
                        begin = children[idx_begin]
                        instr_elements = [e for e in
                                          begin.getparent().findall('{%(w)s}r/{%(w)s}instrText' % NAMESPACES)
                                          if idx_begin < children.index(e.getparent()) < idx_end]
                        if len(instr_elements) == 0:
                            continue

                        # set the text of the first instrText element to the concatenation
                        # of all the instrText element texts
                        instr_text = ''.join([e.text for e in instr_elements])
                        instr_elements[0].text = instr_text

                        # delete all instrText elements except the first
                        for instr in instr_elements[1:]:
                            instr.getparent().remove(instr)

                        name = self.__parse_instr(instr_text)
                        if name is None:
                            continue

                        parent[idx_begin] = Element('MergeField', name=name)

                        # use this so we know *where* to put the replacement
                        instr_elements[0].tag = 'MergeText'
                        block = instr_elements[0].getparent()
                        # append the other tags in the w:r block too
                        parent[idx_begin].extend(list(block))

                        to_delete += [(parent, parent[i + 1])
                                      for i in range(idx_begin, idx_end)]

            for parent, child in to_delete:
                parent.remove(child)

            # Remove mail merge settings to avoid error messages when opening document in Winword
            if self.settings:
                settings_root = self.settings.getroot()
                mail_merge = settings_root.find('{%(w)s}mailMerge' % NAMESPACES)
                if mail_merge is not None:
                    settings_root.remove(mail_merge)
        except:
            self.zip.close()
            raise

    @classmethod
    def __parse_instr(cls, instr):
        args = shlex.split(instr, posix=False)
        if args[0] != 'MERGEFIELD':
            return None
        name = args[1]
        if name[0] == '"' and name[-1] == '"':
            name = name[1:-1]
        return name

    def __get_tree_of_file(self, file):
        fn = file.attrib['PartName' % NAMESPACES].split('/', 1)[1]
        zi = self.zip.getinfo(fn)
        return zi, etree.parse(self.zip.open(zi))

    # def write(self, file):
    #     # Replace all remaining merge fields with empty values
    #     for field in self.get_merge_fields():
    #         self.merge(**{field: ''})

    #     with ZipFile(file, 'w', ZIP_DEFLATED) as output:
    #         for zi in self.zip.filelist:
    #             if zi in self.parts:
    #                 xml = etree.tostring(self.parts[zi].getroot())
    #                 output.writestr(zi.filename, xml)
    #             elif zi == self._settings_info:
    #                 xml = etree.tostring(self.settings.getroot())
    #                 output.writestr(zi.filename, xml)
    #             else:
    #                 output.writestr(zi.filename, self.zip.read(zi))

    def write(self, file):
        # Replace all remaining merge fields with empty values
        for field in self.get_merge_fields():
            self.merge(**{field: ''})

        
        with ZipFile(file, 'w', ZIP_DEFLATED) as output:
            for zi in self.zip.filelist:
                if zi in self.parts:
                    xml = etree.tostring(self.parts[zi].getroot())
                    # if str(zi) == "<ZipInfo filename='word/document.xml' compress_type=deflate file_size=2070956 compress_size=56133>":
                    if 'word/document.xml' in str(zi):
                        # print('true')
                        xml = xml.decode('utf-8')
                        all_table_tags = re.findall('<w:tbl>',xml)
                        # print(all_table_tags)
                        table_start,table_end = 0,len(xml)
                        for i in range(len(all_table_tags)):
                            start_t_index = xml.find('<w:tbl>',table_start,table_end)
                            end_t_index = xml.find('</w:tbl>',start_t_index,table_end)
                            end_t_index += len('</w:tbl>')
                            # print(f"*******\n\n{xml[start_t_index:end_t_index]}\n\n*******")
                            table_xml = xml[start_t_index:end_t_index]
                            all_w_t_tags = re.findall('<w:t>',table_xml)
                            all_empty_w_t_tags = re.findall('<w:t></w:t>',table_xml)
                            # print(f"******\n\n{len(all_w_t_tags), len(all_empty_w_t_tags)}\n\n*********")
                            if len(all_w_t_tags) == len(all_empty_w_t_tags):
                                xml = xml[:start_t_index] + xml[end_t_index:]
                                table_end = len(xml)
                            else:
                                table_start = end_t_index
                        all_empty_tags = re.findall('<w:t></w:t>',xml)
                        total_empty_tags = len(all_empty_tags)
                        # print(total_empty_tags)
                        total_empty_tags_handled = 0
                        start_index=0
                        end_index = len(xml)
                        # print(end_index)
                        if_table_found = False
                        while( total_empty_tags > 0 ):
                            if not if_table_found :
                                start_index = 0
                                end_index = len(xml)
                            table_start_index, table_end_index = 0,len(xml)
                            n = len(xml)
                            all_table_tags_indexses = []
                            for i in range(len(all_table_tags)):
                                # print('finding table tags')
                                table_start_index = xml.find('<w:tbl>',table_start_index,table_end_index)
                                table_end_index = xml.find('</w:tbl>',table_start_index,table_end_index)
                                if table_start_index > -1 and table_end_index > -1:
                                    table_end_index += len('</w:tbl>')
                                    all_table_tags_indexses.append((table_start_index, table_end_index))
                                    # print(xml[table_start_index:table_end_index])
                                    table_start_index = table_end_index
                                    table_end_index = n
                                else:
                                    break

                            ind = xml.find('<w:t></w:t>',start_index,end_index)
                            # start_index = ind + len('<w:t></w:t>')
                            # print(start_index, end_index)
                            is_in_table = False
                            for tags in all_table_tags_indexses:
                                # print('checking if in table')
                                if ind >= tags[0] and ind <= tags[1]:
                                    if_table_found = True
                                    is_in_table = True
                                    # print(tags[0],tags[1])
                                    # print(start_index)
                                    start_index = tags[1]
                                    # print(start_index, end_index)
                                    table = xml[tags[0]:tags[1]]
                                    in_table_empty_tags = re.findall('<w:t></w:t>',table)
                                    # print(total_empty_tags)
                                    total_empty_tags -= len(in_table_empty_tags)
                                    # print(total_empty_tags)
                                    # print('t1:',total_empty_tags ,total_empty_tags_handled)
                                    # print(len(in_table_empty_tags))
                                    # total_empty_tags_handled += len(in_table_empty_tags)
                                    # print('t2:',total_empty_tags ,total_empty_tags_handled)
                                    break
                            
                            if not is_in_table:
                                # print("in here")
                                i = ind
                                while True:
                                    while(xml[i-4:i] != '<w:r' ):
                                        i -= 1
                                    if (xml[i] == ' ' or xml[i] == '>'):
                                        temp_i = i-4
                                        start_i = i-4
                                        end_i = ind+17
                                        break
                                    else:
                                        i -= 1
                                # print(xml[start_i:end_i], end='\n\n')
                                if xml[ind+17 : ind+22] == '<w:r>' or xml[ind+17 : ind+22] == '<w:r ':
                                    end_ind = ind+22
                                    while(xml[end_ind: end_ind+6] != '</w:r>'):
                                        end_ind+=1
                                    temp_end_i = end_ind+6
                                    empty_block = '<w:t xml:space="preserve"> </w:t>'
                                    if xml.find(empty_block, end_i+1,temp_end_i) > 0:
                                        end_i = temp_end_i
                                        if xml[start_i-8:start_i] == '</w:pPr>' and xml[end_i:end_i+6] == '</w:p>':
                                            j = start_i-8
                                            while True:
                                                while(xml[j-4:j] != '<w:p'):
                                                    j = j - 1
                                                if xml[j] == '>' or xml[j] == ' ':
                                                    start_i = j-4
                                                    end_i = end_i + 6
                                                    break
                                                else:
                                                    j -= 1
                                
                                if xml[temp_i-8:temp_i] == '</w:pPr>' and xml[ind+17:ind+23] == '</w:p>':
                                    j = temp_i-8
                                    
                                    while True:
                                        while(xml[j-4:j] != '<w:p'):
                                            j = j - 1
                                        if xml[j] == '>' or xml[j] == ' ':
                                            start_i = j-4
                                            end_i = ind+23
                                            break
                                        else:
                                            j -= 1
                                    # print(xml[start_i:end_i],end='\n\n')

                                # print(len(xml))
                                temp = xml[:start_i] + xml[end_i:]
                                del(xml)
                                xml = temp
                                del(temp)
                                end_index = len(xml)
                                # print(len(xml))
                                # total_empty_tags_handled += 1
                                total_empty_tags -= 1
			
			corrupted_tags = re.findall(r"\d/w:t",xml)
			for i in range(len(corrupted_tags)):
				c_tag_index = xml.find(corrupted_tags[i])
				temp = xml[:c_tag_index] + '<' + xml[c_tag_index+1:]
				del(xml)
				xml = temp
				del(temp)
                        f = open('final_xml.xml','w')
                        f.write(f"{xml}")
                        f.close()
                        xml = bytes(xml,'utf-8')
                    output.writestr(zi.filename, xml)
                elif zi == self._settings_info:
                    xml = etree.tostring(self.settings.getroot())
                    output.writestr(zi.filename, xml)
                else:
                    output.writestr(zi.filename, self.zip.read(zi))


    # def write(self, file):
    #     # Replace all remaining merge fields with empty values
    #     for field in self.get_merge_fields():
    #         self.merge(**{field: ''})


    #     with ZipFile(file, 'w', ZIP_DEFLATED) as output:
    #         for zi in self.zip.filelist:
    #             if zi in self.parts:
    #                 xml = etree.tostring(self.parts[zi].getroot())
    #                 xml = str(xml, 'utf-8')
    #                 while xml.find('<w:t></w:t>') > 0:
    #                     ind = xml.find('<w:t></w:t>')
    #                     i = ind
    #                     while True:
    #                         while(xml[i-4:i] != '<w:r' ):
    #                             i -= 1
    #                         if (xml[i] == ' ' or xml[i] == '>'):
    #                             temp_i = i-4
    #                             start_i = i-4
    #                             end_i = ind+17
    #                             break
    #                         else:
    #                             i -= 1
    #                     # print(xml[start_i:end_i], end='\n\n')
    #                     if xml[ind+17 : ind+22] == '<w:r>' or xml[ind+17 : ind+22] == '<w:r ':
    #                         end_ind = ind+22
    #                         while(xml[end_ind: end_ind+6] != '</w:r>'):
    #                             end_ind+=1
    #                         temp_end_i = end_ind+6
    #                         empty_block = '<w:t xml:space="preserve"> </w:t>'
    #                         if xml.find(empty_block, end_i+1,temp_end_i) > 0:
    #                             end_i = temp_end_i
    #                             if xml[start_i-8:start_i] == '</w:pPr>' and xml[end_i:end_i+6] == '</w:p>':
    #                                 j = start_i-8
    #                                 while True:
    #                                     while(xml[j-4:j] != '<w:p'):
    #                                         j = j - 1
    #                                     if xml[j] == '>' or xml[j] == ' ':
    #                                         start_i = j-4
    #                                         end_i = end_i + 6
    #                                         break
    #                                     else:
    #                                         j -= 1
                        
    #                     if xml[temp_i-8:temp_i] == '</w:pPr>' and xml[ind+17:ind+23] == '</w:p>':
    #                         j = temp_i-8
                            
    #                         while True:
    #                             while(xml[j-4:j] != '<w:p'):
    #                                 j = j - 1
    #                             if xml[j] == '>' or xml[j] == ' ':
    #                                 start_i = j-4
    #                                 end_i = ind+23
    #                                 break
    #                             else:
    #                                 j -= 1
    #                         # print(xml[start_i:end_i],end='\n\n')

    #                     temp = xml[:start_i] + xml[end_i:]
    #                     del(xml)
    #                     xml = temp
    #                     del(temp)
    #                 xml = bytes(xml,'utf-8')
    #                 output.writestr(zi.filename, xml)
    #             elif zi == self._settings_info:
    #                 xml = etree.tostring(self.settings.getroot())
    #                 xml2 = str(xml, 'utf-8')
    #                 while xml2.find('<w:t></w:t>') > 0:
    #                     ind = xml2.find('<w:t></w:t>')
    #                     i = ind
    #                     while True:
    #                         while(xml[i-4:i] != '<w:r' ):
    #                             i -= 1
    #                         if (xml[i] == '>' or xml[i] == ' '):
    #                             temp_i = i-4
    #                             start_i = i-4
    #                             end_i = ind+17
    #                             break
    #                         else:
    #                             i -=1
    #                     # print(xml[start_i:end_i],end='\n\n')
    #                     if xml2[ind+17 : ind+22] == '<w:r>' or xml2[ind+17 : ind+22] == '<w:r ':
    #                         end_ind = ind+22
    #                         while(xml2[end_ind: end_ind+6] != '</w:r>'):
    #                             end_ind+=1
    #                         temp_end_i = end_ind+6
    #                         empty_block = '<w:t xml:space="preserve"> </w:t>'
    #                         if xml2.find(empty_block, end_i+1,temp_end_i) > 0:
    #                             end_i = temp_end_i
    #                             if xml2[start_i-8:start_i] == '</w:pPr>' and xml2[end_i:end_i+6] == '</w:p>':
    #                                 j = start_i-8
    #                                 while True:
    #                                     while(xml2[j-4:j] != '<w:p'):
    #                                         j = j - 1
    #                                     if xml2[j] == '>' or xml2[j] == ' ':
    #                                         start_i = j-4
    #                                         end_i = end_i+6
    #                                         break
    #                                     else:
    #                                         j -= 1
                        
    #                     if xml2[temp_i-8:temp_i] == '</w:pPr>' and xml2[ind+17:ind+23] == '</w:p>':
    #                         j = temp_i-8
    #                         # while(xml2[j-4:j] != '<w:p'):
    #                         #     j = j - 1
    #                         # start_i = j-5
    #                         # end_i = ind+23
    #                         while True:
    #                             while(xml2[j-4:j] != '<w:p'):
    #                                 j = j - 1
    #                             if xml2[j] == '>' or xml2[j] == ' ':
    #                                 start_i = j-4
    #                                 end_i = ind+23
    #                                 break
    #                             else:
    #                                 j -= 1
    #                         # print(xml[start_i:end_i],end='\n\n')

    #                     temp2 = xml2[:start_i] + xml2[end_i:]
    #                     del(xml2)
    #                     xml2 = temp2
    #                     del(temp2)
    #                 xml = bytes(xml2,'utf-8')
    #                 output.writestr(zi.filename, xml)
    #             else:
    #                 output.writestr(zi.filename, self.zip.read(zi))

    def get_merge_fields(self, parts=None):
        if not parts:
            parts = self.parts.values()
        fields = set()
        for part in parts:
            for mf in part.findall('.//MergeField'):
                fields.add(mf.attrib['name'])
        return fields

    def merge_templates(self, replacements, separator):
        """
        Duplicate template. Creates a copy of the template, does a merge, and separates them by a new paragraph, a new break or a new section break.
        separator must be :
        - page_break : Page Break. 
        - column_break : Column Break. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - textWrapping_break : Line Break.
        - continuous_section : Continuous section break. Begins the section on the next paragraph.
        - evenPage_section : evenPage section break. section begins on the next even-numbered page, leaving the next odd page blank if necessary.
        - nextColumn_section : nextColumn section break. section begins on the following column on the page. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - nextPage_section : nextPage section break. section begins on the following page.
        - oddPage_section : oddPage section break. section begins on the next odd-numbered page, leaving the next even page blank if necessary.
        """

        #TYPE PARAM CONTROL AND SPLIT
        valid_separators = {'page_break', 'column_break', 'textWrapping_break', 'continuous_section', 'evenPage_section', 'nextColumn_section', 'nextPage_section', 'oddPage_section'}
        if not separator in valid_separators:
            raise ValueError("Invalid separator argument")
        type, sepClass = separator.split("_")
  

        #GET ROOT - WORK WITH DOCUMENT
        for part in self.parts.values():
            root = part.getroot()
            tag = root.tag
            if tag == '{%(w)s}ftr' % NAMESPACES or tag == '{%(w)s}hdr' % NAMESPACES:
                continue
		
            if sepClass == 'section':

                #FINDING FIRST SECTION OF THE DOCUMENT
                firstSection = root.find("w:body/w:p/w:pPr/w:sectPr", namespaces=NAMESPACES)
                if firstSection == None:
                    firstSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)
			
                #MODIFY TYPE ATTRIBUTE OF FIRST SECTION FOR MERGING
                nextPageSec = deepcopy(firstSection)
                for child in nextPageSec:
                #Delete old type if exist
                    if child.tag == '{%(w)s}type' % NAMESPACES:
                        nextPageSec.remove(child)
                #Create new type (def parameter)
                newType = etree.SubElement(nextPageSec, '{%(w)s}type'  % NAMESPACES)
                newType.set('{%(w)s}val'  % NAMESPACES, type)

                #REPLACING FIRST SECTION
                secRoot = firstSection.getparent()
                secRoot.replace(firstSection, nextPageSec)

            #FINDING LAST SECTION OF THE DOCUMENT
            lastSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)

            #SAVING LAST SECTION
            mainSection = deepcopy(lastSection)
            lsecRoot = lastSection.getparent()
            lsecRoot.remove(lastSection)

            #COPY CHILDREN ELEMENTS OF BODY IN A LIST
            childrenList = root.findall('w:body/*', namespaces=NAMESPACES)

            #DELETE ALL CHILDREN OF BODY
            for child in root:
                if child.tag == '{%(w)s}body' % NAMESPACES:
                    child.clear()

            #REFILL BODY AND MERGE DOCS - ADD LAST SECTION ENCAPSULATED OR NOT
            lr = len(replacements)
            lc = len(childrenList)

            for i, repl in enumerate(replacements):
                parts = []
                for (j, n) in enumerate(childrenList):
                    element = deepcopy(n)
                    for child in root:
                        if child.tag == '{%(w)s}body' % NAMESPACES:
                            child.append(element)
                            parts.append(element)
                            if (j + 1) == lc:
                                if (i + 1) == lr:
                                    child.append(mainSection)
                                    parts.append(mainSection)
                                else:
                                    if sepClass == 'section':
                                        intSection = deepcopy(mainSection)
                                        p   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        pPr = etree.SubElement(p, '{%(w)s}pPr'  % NAMESPACES)
                                        pPr.append(intSection)
                                        parts.append(p)
                                    elif sepClass == 'break':
                                        pb   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        r = etree.SubElement(pb, '{%(w)s}r'  % NAMESPACES)
                                        nbreak = Element('{%(w)s}br' % NAMESPACES)
                                        nbreak.attrib['{%(w)s}type' % NAMESPACES] = type
                                        r.append(nbreak)

                    self.merge(parts, **repl)

    def merge_pages(self, replacements):
         """
         Deprecated method.
         """
         warnings.warn("merge_pages has been deprecated in favour of merge_templates",
                      category=DeprecationWarning,
                      stacklevel=2)         
         self.merge_templates(replacements, "page_break")

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for field, replacement in replacements.items():
            if isinstance(replacement, list):
                self.merge_rows(field, replacement)
            else:
                for part in parts:
                    self.__merge_field(part, field, replacement)

    def __merge_field(self, part, field, text):
        for mf in part.findall('.//MergeField[@name="%s"]' % field):
            children = list(mf)
            mf.clear()  # clear away the attributes
            mf.tag = '{%(w)s}r' % NAMESPACES
            mf.extend(children)

            nodes = []
            # preserve new lines in replacement text
            text = text or ''  # text might be None
            text_parts = str(text).replace('\r', '').split('\n')
            for i, text_part in enumerate(text_parts):
                text_node = Element('{%(w)s}t' % NAMESPACES)
                text_node.text = text_part
                nodes.append(text_node)

                # if not last node add new line node
                if i < (len(text_parts) - 1):
                    nodes.append(Element('{%(w)s}br' % NAMESPACES))

            ph = mf.find('MergeText')
            if ph is not None:
                # add text nodes at the exact position where
                # MergeText was found
                index = mf.index(ph)
                for node in reversed(nodes):
                    mf.insert(index, node)
                mf.remove(ph)
            else:
                mf.extend(nodes)

    def merge_rows(self, anchor, rows):
        table, idx, template = self.__find_row_anchor(anchor)
        if table is not None:
            if len(rows) > 0:
                del table[idx]
                for i, row_data in enumerate(rows):
                    row = deepcopy(template)
                    self.merge([row], **row_data)
                    table.insert(idx + i, row)
            else:
                # if there is no data for a given table
                # we check whether table needs to be removed
                if self.remove_empty_tables:
                    parent = table.getparent()
                    parent.remove(table)

    def __find_row_anchor(self, field, parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            for table in part.findall('.//{%(w)s}tbl' % NAMESPACES):
                for idx, row in enumerate(table):
                    if row.find('.//MergeField[@name="%s"]' % field) is not None:
                        return table, idx, row
        return None, None, None

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def close(self):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None
