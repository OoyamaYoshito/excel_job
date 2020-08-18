def wsgi_app(environ, start_response):
    import sys
    import pprint
    import os
    sys.path.append('/Users/member/git/excel_job')
    import output
    #output = pprint.pformat(os.listdir('.')) + '\n'
    output = sys.version + '\n' 
    for q in environ['QUERY_STRING'].split('&'):
        if q.startswith('grade='):
            stgrade=q.replace('grade=','')
        if q.startswith('class='):
            stclass=q.replace('class=','')
    if stgrade != '' and stclass != '':
        output += 'grade=' + stgrade + '\n'
        output += 'class=' + stclass + '\n'
    output = output.encode('utf8')
    status = '200 OK'
    headers = [('Content-type', 'text/plain'),
               ('Content-Length', str(len(output)))]
    start_response(status, headers)
    yield output

application = wsgi_app
