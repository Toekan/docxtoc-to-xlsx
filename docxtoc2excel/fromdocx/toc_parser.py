import re


def import_and_tokenize_toc(doc):
    list_content_toc = import_toc(doc)
    list_tokenized_toc = tokenize_toc(list_content_toc)
    return list_tokenized_toc


def import_toc(doc):
    """"""
    table_of_content = [paragraph for paragraph in doc.paragraphs if
                        paragraph.style.name[:3] == 'toc']

    lines_tokenized = []
    for paragraph in table_of_content:
        line_tokenized = [run.text for run in paragraph.runs
                   if run.text not in ['\t', '']][:-1]
        lines_tokenized.append(line_tokenized)
    return lines_tokenized


def tokenize_toc(lines_tokenized):
    """"""
    regex = re.compile('\|\w*\|[\w\s]*$')
    regex_description = re.compile('\|\w*\|')
    lines_tokenized_updated = []
    for i in range(len(lines_tokenized)):

        if re.search(regex, lines_tokenized[i][-1]):
            if len(lines_tokenized[i]) > 3:
                s = ''
                for element in lines_tokenized[i][1:-1]:
                    s += element
                line = [lines_tokenized[i][0], s, lines_tokenized[i][-1]]
            else:
                line = lines_tokenized[i][:]

            match = re.findall(regex, line[-1])[0]
            description = re.findall(regex_description, match)[0]
            dimension = match.replace(description, '')
            line[-1] = description
            line.append(dimension)
            lines_tokenized_updated.append(line)
        else:
            if len(lines_tokenized[i]) > 2:
                s = ''
                for element in lines_tokenized[i][1:]:
                    s += element
                line = [lines_tokenized[i][0], s]
            else:
                line = lines_tokenized[i]
            lines_tokenized_updated.append(line)
    return lines_tokenized_updated
