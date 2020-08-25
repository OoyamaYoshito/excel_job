outpath="/Users/member/git/excel_job/test"

def output_excel(stgrade, stclass, outpath):
    import output
    student="/Users/member/git/excel_job/studentlist"
    answer="/Users/member/git/excel_job/answersdata"
    import re
    import os
    basepath=re.sub('/[^/]+$','/',__file__)
    os.chdir(basepath)
    for g in stgrade:
        for c in stclass:
            output.output_class(g,c,outputpath=outpath,studentpath=student,answerpath=answer)

def wsgi_app(environ, start_response):
    import sys
    import pprint
    import os
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
        output += 'grade=' + stgrade + '\n'
        output += 'class=' + stclass + '\n'
        output_excel(stgrade,stclass,outpath)
    else:
        output = '<form>'
        output += 'Grade: '
        for g in [1,2,3,4]:
            output += '<input type="radio" name="grade" value="' + str(g) + '">' + str(g)
        output += '<br/>'
        output += 'Class: '
        for c in 'ABCDEFGHIJKL':
            output += '<input type="checkbox" name="class" value="' + str(c) + '">' + str(c)
        output += '<br/>'
        output += '<input type="submit" value="send"/>'
        output += '</form>'
    output = output.encode('utf-8')
    status = '200 OK'
    headers = [('Content-type', 'text/html'),
               ('Content-Length', str(len(output)))]
    start_response(status, headers)
    yield output

application = wsgi_app
