#outpath="/Users/member/git/excel_job/test"

def output_excel(stgrade, stclass, outpath):
    import output
    student="/Users/member/git/excel_job/studentlist"
    answer="/Users/member/git/excel_job/answersdata"
    import os
    os.chdir(os.path.dirname(__file__))
    targets=[]
    try:
        for g in stgrade:
            for c in stclass:
                output.output_class(g,c,outputpath=outpath,studentpath=student,answerpath=answer)
                targets.append('output' + g + c + '.xlsx')
    except ValueError as e:
        return e.args[0]
    import zipfile
    zipname=outpath+'/output'+stgrade+stclass+'.zip'
    with zipfile.ZipFile(zipname,'w',compression=zipfile.ZIP_DEFLATED) as zip:
        for f in targets:
            zip.write(outpath + '/' + f, arcname=f)
    return zipname

def output_excel2(students, outpath):
    import output2
    student="/Users/member/git/excel_job/studentlist"
    answer="/Users/member/git/excel_job/answersdata"
    import os
    os.chdir(os.path.dirname(__file__))
    try:
        output2.output_students(students,outputpath=outpath,studentpath=student,answerpath=answer)
    except ValueError as e:
        return e.args[0]
    return outpath + '/output.xlsx'

def wsgi_app(environ, start_response):
    import sys
    import pprint
    import os
    import re
    import urllib
    sys.path.append('/Users/member/git/excel_job')
    import output
    #output = pprint.pformat(os.listdir('.')) + '\n'
    output = sys.version + '\n' 
    select = ''
    stgrade = ''
    stclass = ''
    students = ''
    for q in environ['QUERY_STRING'].split('&'):
        if q.startswith('select='):
            select=q.replace('select=','')
        if q.startswith('grade='):
            stgrade=q.replace('grade=','')
        if q.startswith('class='):
            stclass+=q.replace('class=','')
        if q.startswith('students='):
            students+=q.replace('students=','')
    if select == 'byclass' and stgrade != '' and stclass != '':
        import tempfile
        outpath=tempfile.mkdtemp()
        zip = output_excel(stgrade,stclass,outpath)
        if os.path.isfile(zip):
            headers = [('Content-type', 'application/zip'),
                       ('Content-Length', str(os.path.getsize(zip))),
                       ('Content-Disposition', 'attachment; filename="' + os.path.basename(zip) + '"')]
            with open(zip,'rb') as f:
                output=f.read()
                f.close()
        else:
            output = zip
            output = output.encode('utf-8')
            headers = [('Content-type', 'text/html; charset=utf-8'),
                       ('Content-Length', str(len(output)))]
    elif select == 'bystudents' and students != '':
        output = urllib.parse.unquote(students)
        students = output.split()
        import tempfile
        outpath=tempfile.mkdtemp()
        zip = output_excel2(students,outpath)
        if os.path.isfile(zip):
            headers = [('Content-type', 'application/excel'),
                       ('Content-Length', str(os.path.getsize(zip))),
                       ('Content-Disposition', 'attachment; filename="' + os.path.basename(zip) + '"')]
            with open(zip,'rb') as f:
                output=f.read()
                f.close()
        else:
            output = zip
            output = output.encode('utf-8')
            headers = [('Content-type', 'text/html; charset=utf-8'),
                       ('Content-Length', str(len(output)))]
    else:
        output = '<script type="text/javascript">'
        output += 'function byclass() {'
        output += '  var x = document.getElementsByName("grade");'
        output += '  for (i = 0; i < x.length; i++) x[i].disabled = false;'
        output += '  var x = document.getElementsByName("class");'
        output += '  for (i = 0; i < x.length; i++) x[i].disabled = false;'
        output += '  document.getElementsByName("students")[0].disabled = true;'
        output += '}'
        output += 'function bystudents() {'
        output += '  var x = document.getElementsByName("grade");'
        output += '  for (i = 0; i < x.length; i++) x[i].disabled = true;'
        output += '  var x = document.getElementsByName("class");'
        output += '  for (i = 0; i < x.length; i++) x[i].disabled = true;'
        output += '  document.getElementsByName("students")[0].disabled = false;'
        output += '}'
        output += '</script>'
        output += '<form>'
        output += 'Select: '
        output += '<input type="radio" name="select" value="byclass" checked onclick="byclass()">By Class'
        output += '<input type="radio" name="select" value="bystudents" onclick="bystudents()">By Student IDs'
        output += '<br/>'
        output += 'Grade: '
        for g in [1,2]:
            output += '<input type="radio" name="grade" value="' + str(g) + '">' + str(g)
        output += '<br/>'
        output += 'Class: '
        for c in 'ABCDEFGHIJKL':
            output += '<input type="checkbox" name="class" value="' + str(c) + '">' + str(c)
        output += '<br/>'
        output += 'Students: <textarea name="students" disabled></textarea>'
        output += '<br/>'
        output += '<input type="submit" value="send"/>'
        output += '</form>'
        output = output.encode('utf-8')
        headers = [('Content-type', 'text/html; charset=utf-8'),
                   ('Content-Length', str(len(output)))]
    status = '200 OK'
    start_response(status, headers)
    yield output

application = wsgi_app
