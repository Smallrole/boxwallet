import sys
from gnosis.xml.objectify import XML_Objectify, pyobj_printer, EXPAT, DOM
LF = "\n"
def show(xml_src, parser):
    """Self test using simple or user-specified XML data
    >>> xml = '''
    ...
    ...
    ...    Some text about eggs.
    ...    Ode to Spam
    ...  '''
    >>> squeeze = lambda s: s.replace(LF*2,LF).strip()
    >>> print squeeze(show(xml,DOM)[0])
    -----* _XO_Spam *-----
    {Eggs}
       PCDATA=Some text about eggs.
    {MoreSpam}
       PCDATA=Ode to Spam
    >>> print squeeze(show(xml,EXPAT)[0])
    -----* _XO_Spam *-----
    {Eggs}
       PCDATA=Some text about eggs.
    {MoreSpam}
       PCDATA=Ode to Spam
    PCDATA=
    """
    try:
        xml_obj = XML_Objectify(xml_src, parser=parser)
        py_obj = xml_obj.make_instance()
        return (pyobj_printer(py_obj).encode('UTF-8'),
                "++ SUCCESS (using "+parser+")\n")
    except:
        return ("","++ FAILED (using "+parser+")\n")
if __name__ == "__main__":
    if len(sys.argv)==1 or sys.argv[1]=="-v":
        import doctest, test_basic
        doctest.testmod(test_basic,verbose=True)
    elif sys.argv[1] in ('-h','-help','--help'):
        print "You may specify XML files to objectify instead of self-test"
        print "(Use '-v' for verbose output, otherwise no message means success)"
    else:
        for filename in sys.argv[1:]:
            for parser in (DOM, EXPAT):
                output, message = show(filename, parser)
                print output
                sys.stderr.write(message)
                print "=" * 50
