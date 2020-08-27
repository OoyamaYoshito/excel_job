outpath="/Users/member/git/excel_job/test"

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

def wsgi_app(environ, start_response):
    import sys
    import pprint
    import os
    import re
    sys.path.append('/Users/member/git/excel_job')
    import output
    #output = pprint.pformat(os.listdir('.')) + '\n'
    output = sys.version + '\n' 
    stgrade = ''
    stclass = ''
    for q in environ['QUERY_STRING'].split('&'):
        if q.startswith('grade='):
            stgrade=q.replace('grade=','')
        if q.startswith('class='):
            stclass+=q.replace('class=','')
    if stgrade != '' and stclass != '':
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
    else:
        output = '<form>'
        output += 'Grade: '
        for g in [1,2]:
            output += '<input type="radio" name="grade" value="' + str(g) + '">' + str(g)
        output += '<br/>'
        output += 'Class: '
        for c in 'ABCDEFGHIJKL':
            output += '<input type="checkbox" name="class" value="' + str(c) + '">' + str(c)
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
