#!/usr/bin/env python3

import os, sys, pprint, collections, html, string

from html.parser import HTMLParser
from xml.dom import minidom

import PyPDF2

LATEX_CONVERSIONS = [('&', '\\&'), ('#', '\\#'), ('_', '\\_'), ('<p>', ''), ('</p>', '\n\n'), ('<strong>', '\\textbf{'), ('</strong>', '}'), ('<em>', '\\textit{'), ('</em>', '}'), ]

def TransformToDict(node):
    #print(dir(node))
    N = node.nodeName
    if N == '#text':
        data = node.data.strip()
        if data:
            return {N: data}
        else:
            return None
    else:
        L = []
        for e in node.childNodes:
            w = TransformToDict(e)
            if w is not None:
                L.append(w)
        if node.hasAttributes():
            A = []
            for i in range(node.attributes.length):
                a = node.attributes.item(i)
                #print("attr {} {}".format(a.name, a.value))
                A.append((a.name, a.value))
            L.append({"#attributes": A})
        return {N: L}

def FindKey(D, key):
    Q = collections.deque()
    if type(D) is dict:
        Q.append(D)
    elif type(D) is list:
        Q.extend(D)
    else:
        assert False, 'unknown type for D: ' + type(D)
    while len(Q) > 0:
        front = Q.popleft()
        assert type(front) is dict
        if key in front:
            return front[key]
        for x in front:
            if type(front[x]) is list:
                for e in front[x]:
                    if type(e) is dict:
                        Q.append(e)
    raise ValueError('key not found: ' + key)

def GetText(D, key):
    if type(D) is not list:
        D = [D]
    for x in D:
        try:
            node = FindKey(x, key)
        except ValueError:
            continue 

        for e in node:
            text = FindKey(e, '#text')
            return text 

class TextWrap:
    def __init__(self, D):
        self.D = D 

    def __getitem__(self, key):
        return GetText(self.D, key)

class MyHTMLParser(HTMLParser):
    def __init__(self):
        super(MyHTMLParser, self).__init__()
        self.latex = ''

    def handle_starttag(self, tag, attrs):
        attrs_dict = dict(attrs)
        #print("Encountered a start tag:", tag)
        if tag.lower() == 'a':
            self.latex += '\\href{' + attrs_dict.get('href', '') + '}{'
        elif tag.lower() == 'ul':
            self.latex += '\\begin{itemize}\\setlength\\itemsep{0em}\n'
        elif tag.lower() == 'li':
            self.latex += '\\item '
        else:
            assert False

    def handle_endtag(self, tag):
        #print("Encountered an end tag :", tag)
        if tag.lower() == 'a':
            self.latex += '}'
        elif tag.lower() == 'ul':
            self.latex += '\\end{itemize}\n'
        elif tag.lower() == 'li':
            self.latex += '\n'
        else:
            assert False

    def handle_data(self, data):
        #print("Encountered some data  :", data)
        self.latex += data

    def to_latex(self, data):
        self.feed(data)
        return self.latex

def LatexText(S):
    if S is None:
        return ''

    S = html.unescape(S)
    T = ''

    special_chars = '\'"<>{}[]'

    i = 0
    while i < len(S):
        while i < len(S) and S[i] in string.whitespace:
            T += S[i]
            i = i + 1
        if i < len(S):
            if (S.startswith('http://', i) or S.startswith('https://', i)) and (i == 0 or S[i - 1] not in special_chars):
                j = i 
                while i < len(S) and S[i] not in (string.whitespace + special_chars):
                    i = i + 1
                #print(S[j:i])
                T += '<a href="{0}">{0}</a>'.format(S[j:i])
            else:
                T += S[i]
                i = i + 1

    S = T

    for k, v in LATEX_CONVERSIONS:
        S = S.replace(k, v)

    parser = MyHTMLParser()
    S = parser.to_latex(S)

    return S

def ConvertPeriod(PeriodNode):
    PeriodDict = dict(PeriodNode[0]['#attributes'])
    PeriodStr = PeriodDict.get('year', '').replace('-', '')
    if 'month' in PeriodDict:
        PeriodStr += '/' + PeriodDict.get('month', '').replace('-', '')
        if 'day' in PeriodDict:
            PeriodStr += '/' + PeriodDict.get('day', '').replace('-', '')
    return PeriodStr

def ConvertPeriodToString(Period):
    try:
        PeriodFromStr = ConvertPeriod(FindKey(Period, 'From'))
    except ValueError:
        PeriodFromStr = ''

    try:
        PeriodToStr = ConvertPeriod(FindKey(Period, 'To'))
    except ValueError:
        PeriodToStr = ''

    if PeriodFromStr and PeriodToStr:
        WorkPeriod = PeriodFromStr + ' -- ' + PeriodToStr
    elif PeriodFromStr:
        WorkPeriod = PeriodFromStr
    elif PeriodToStr:
        WorkPeriod = PeriodToStr
    else:
        WorkPeriod = ''

    return WorkPeriod

def ConvertEmployerAddressToString(EmployerNode):
    EmployerMunicipality = LatexText(GetText(EmployerNode, 'Municipality'))
    EmployerLabel = LatexText(GetText(EmployerNode, 'Label'))
    
    if EmployerMunicipality:
        EmployerAddr = ", {}, ({})".format(EmployerMunicipality, EmployerLabel)
    elif EmployerLabel:
        EmployerAddr = ", ({})".format(EmployerLabel)
    else:
        EmployerAddr = ''

    return EmployerAddr

def Visit(D):
    assert type(D) is dict 
    assert len(D.keys()) == 1

    Identification = FindKey(D, 'Identification')
    TelephoneList = FindKey(Identification, 'TelephoneList')
    WebsiteList = FindKey(Identification, 'WebsiteList')
    InstantMessagingList = FindKey(Identification, 'InstantMessagingList')

    Header = {}
    for x in ['FirstName', 'Surname', 'AddressLine', 'PostalCode', 'Municipality', 'Email']:
        Header[x] = GetText(Identification, x)

    Header['Country'] = GetText(FindKey(Identification, 'Country'), 'Label')
    Header['CountryCode'] = GetText(FindKey(Identification, 'Country'), 'Code')
    Header['TelephoneList'] = [(GetText(x, 'Label'), GetText(x, 'Contact')) for x in TelephoneList]
    Header['WebsiteList'] = [GetText(x, 'Contact') for x in WebsiteList]
    Header['InstantMessagingList'] = [(GetText(x, 'Label'), GetText(x, 'Contact')) for x in InstantMessagingList]

    WebsiteListTex = []

    for x in Header['WebsiteList']:
        if 'linkedin.com' in x.lower():
            WebsiteListTex.append("\\linkedin{{{0}}}".format(x))
        else:
            WebsiteListTex.append("\\personalweb{{{0}}}".format(x))

    TelephoneListTex = []

    for k, v in Header['TelephoneList']:
        if k.lower() == 'mobile':
            TelephoneListTex.append("\\phone{{{0}}}".format(v))
        else:
            TelephoneListTex.append("\\phone{{{0}}}".format(v))

    Header['WebsiteListTex'] = '\\\\'.join(WebsiteListTex)
    Header['TelephoneListTex'] = '\\\\'.join(TelephoneListTex)

    Headline = FindKey(D, 'Headline')

    HeadlineName = LatexText(GetText(FindKey(Headline, 'Type'), 'Label'))
    HeadlineText = LatexText(GetText(FindKey(Headline, 'Description'), 'Label'))

    Header['HeadlineName'] = HeadlineName
    Header['HeadlineText'] = HeadlineText

    WorkExperienceList = FindKey(D, 'WorkExperienceList')
    EducationList = FindKey(D, 'EducationList')
    AchievementList = FindKey(D, 'AchievementList')

    #pprint.pprint(Header)

    tex_Header = """

\\documentclass[a4paper,11pt]{{article}}
%\\usepackage[utf8]{{inputenc}}

\\usepackage{{lmodern}}% http://ctan.org/pkg/lm
\\usepackage{{anyfontsize}}
\\usepackage{{textcomp}}
\\renewcommand{{\\familydefault}}{{\\sfdefault}}

\\usepackage{{sectsty}}
\\usepackage{{xcolor}}

\\sectionfont{{\\color[HTML]{{0e4194}}}}
\\subsectionfont{{\\color[HTML]{{0e4194}}}}

%\\usepackage{{fontspec}}
%\\usepackage{{newunicodechar}}
\\usepackage[english]{{babel}}
%\\usepackage{{libertine}}


\\usepackage{{hyperref}}
\\usepackage[cm]{{fullpage}}
%\\usepackage[margin=1cm]{{geometry}}

\\usepackage{{ifthen}}
\\usepackage{{multicol}}
\\usepackage{{graphicx}}
\\usepackage{{subcaption}}

%opening

\\newcommand{{\\name}}[1]{{{{\\Huge{{}}#1}}}}
\\newcommand{{\\address}}[1]{{{{\\large{{}}#1}}}}
\\newcommand{{\\phone}}[1]{{Phone: \\texttt{{#1}}}}
\\newcommand{{\\email}}[1]{{Email: \\texttt{{#1}}}}
\\newcommand{{\\skype}}[1]{{Skype: \\texttt{{#1}}}}
\\newcommand{{\\linkedin}}[1]{{LinkedIn: \\url{{#1}}}}
\\newcommand{{\\personalweb}}[1]{{Website: \\url{{#1}}}}

\\renewcommand{{\\abstractname}}{{}}

\\usepackage{{fancyhdr}}
\\setlength{{\\headheight}}{{15.2pt}}
\\pagestyle{{fancy}}
\\renewcommand{{\\headrulewidth}}{{0pt}}
%\\fancyhf{{}}
\\rhead{{\\ifthenelse{{\\value{{page}}=1}}{{}}{{{FirstName} {Surname}}}}}
\\lhead{{\\ifthenelse{{\\value{{page}}=1}}{{}}{{Curriculum vitae}}}}
%\\rfoot{{}}


\\begin{{document}}

\\begin{{center}}
\\name{{{FirstName} {Surname}}}\\\\%
\\address{{{AddressLine}, {PostalCode}, {Municipality} {Country} ({CountryCode})}}\\\\%
%\\phone{{{TelephoneList[0][0]}: {TelephoneList[0][1]}}}
{TelephoneListTex} // \\email{{{Email}}} // \\skype{{{InstantMessagingList[0][1]}}}\\\\%
%\\linkedin{{ }}\\\\%
%\\personalweb{{ }}
{WebsiteListTex}
\\end{{center}}

\\section{{{HeadlineName} }}

{HeadlineText}
""".format(**Header)

    tex_Footer = """
%% FOOTER %%

\\input{{alejandro_santos_electronics}}

\\setcounter{{tocdepth}}{{1}}
\\tableofcontents

\\end{{document}}

""".format()

    tex_WorkExperience = """\\section{{Work Experience}}
""".format()

    for WorkNode in WorkExperienceList:
        WorkExp = {}

        WorkExp['Position'] = LatexText(GetText(WorkNode, 'Position'))
        WorkExp['Activities'] = LatexText(GetText(WorkNode, 'Activities'))
        WorkExp['EmployerName'] = LatexText(GetText(FindKey(WorkNode, 'Employer'), 'Name'))
        WorkExp['WorkPeriod'] = ConvertPeriodToString(FindKey(WorkNode, 'Period'))
        WorkExp['EmployerAddr'] = ConvertEmployerAddressToString(FindKey(WorkNode, 'Employer'))

        tex_WorkExperience += """\\subsection{{{WorkPeriod} {Position}}}

\\textbf{{{EmployerName}{EmployerAddr}}}

%\\noindent{{}}
{Activities}
""".format(**WorkExp)

    tex_Education = """\\section{{Education and Training}}
""".format()

    for EduNode in EducationList:
        EduExp = {}

        EduExp['Title'] = LatexText(GetText(EduNode, 'Title'))
        EduExp['Activities'] = LatexText(GetText(EduNode, 'Activities'))
        EduExp['OrgName'] = LatexText(GetText(FindKey(EduNode, 'Organisation'), 'Name'))
        EduExp['WorkPeriod'] = ConvertPeriodToString(FindKey(EduNode, 'Period'))
        EduExp['OrgAddr'] = ConvertEmployerAddressToString(FindKey(EduNode, 'Organisation'))

        tex_Education += """\\subsection{{{WorkPeriod} {Title}}}

\\textbf{{{OrgName}{OrgAddr}}}

%\\noindent{{}}
{Activities}
""".format(**EduExp)

    Skills = {}
    Skills['JobRelated'] = LatexText(GetText(FindKey(D, 'Skills'), 'JobRelated'))

    tex_Skills = """\\section{{Skills: Job related}}

{JobRelated}
""".format(**Skills)

    tex_Achievements = """\\section{{Additional Information}}
""".format()

    for AchNode in AchievementList:
        AchExp = {}

        AchExp['Title'] = LatexText(GetText(FindKey(AchNode, 'Title'), 'Label'))
        AchExp['Description'] = LatexText(GetText(AchNode, 'Description'))

        tex_Achievements += """\\subsection{{{Title}}}

{Description}
""".format(**AchExp)

    print(tex_Header)
    print(tex_Skills)
    print(tex_Education)
    print(tex_WorkExperience)
    print(tex_Achievements)
    print(tex_Footer)

def getAttachments(reader):
      """
      Retrieves the file attachments of the PDF as a dictionary of file names
      and the file data as a bytestring.
      :return: dictionary of filenames and bytestrings
      """
      catalog = reader.trailer["/Root"]
      fileNames = catalog['/Names']['/EmbeddedFiles']['/Names']
      attachments = {}
      for f in fileNames:
          if isinstance(f, str):
              name = f
              dataIndex = fileNames.index(f) + 1
              fDict = fileNames[dataIndex].getObject()
              fData = fDict['/EF']['/F'].getData()
              attachments[name] = fData

      return attachments

def main(argv):
    if argv[1].lower().endswith('.xml'):
        dom = minidom.parse(argv[1])
    elif argv[1].lower().endswith('.pdf'):
        with open(argv[1], 'rb') as f:
            reader = PyPDF2.PdfFileReader(f)
            attchs = getAttachments(reader)
        assert len((attchs)) == 1
        dom = minidom.parseString(list(attchs.items())[0][1]) 
    else:
        assert False

    assert len(dom.childNodes) == 1
    assert dom.childNodes[0].nodeName == 'SkillsPassport'

    D = TransformToDict(dom.childNodes[0])

    #pprint.pprint(D, compact=False, width=900)

    Visit(D)


if __name__ == '__main__':
    main(sys.argv)
