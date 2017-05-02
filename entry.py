# coding=utf-8
'''
https://bitbucket.org/decalage/oletools/downloads
https://bitbucket.org/decalage/oletools/wiki/olevba
pip install oletools
'''

# ------------------------------------------------------------------------------
# CHANGELOG:
# 2016-07-18 v1.00  first version
# 2016-07-26 v1.01  修复宏名字 UnicodeEncodeError错误
# 2016-08-12 v1.10  嵌入文件解包优化
# 2016-08-12 v1.11  文件名字错误
# 2016-08-19 v1.12  eml 附件文件名字非法处理
# 2016-11-19 v1.30  把文件名规范化的代码孤立出来供其它函数使用，多个位置可以解出内嵌文件，内嵌文件名字不可靠
# 2017-01-19 v2.00  彻底解决文件名字的问题，独立模块 io_in_out.py
# 2017-05-02 v3.00  重写


import argparse
import os
import sys

curpath = os.path.dirname(os.path.realpath(__file__))
sys.path.append(os.path.abspath(os.path.join(curpath, '../')))
from io_in_out import *

curpath = io_in_arg(curpath)


def dump_sub_file(host_fullpath, filename_from_host, data_or_fileobj_to_write):
    '''
    从文件中内嵌出来的文件可能文件名是无效的，无法创建文件，这个函数来规范化文件名
    
    :return: the final sub file fullpath 
    '''

    import random
    import shutil

    # must detect path sep first
    _func_replace_os_path_sep = lambda x: x.replace(u'/', u'_').replace(u'\\', u'_')
    filename_from_host = _func_replace_os_path_sep(filename_from_host)

    sub_fullpath = u'{0}_{1}'.format(host_fullpath, filename_from_host)

    if not io_is_path_valid(sub_fullpath):
        exts = os.path.splitext(sub_fullpath)
        ext = exts[1]
        ext = ext if ext and io_is_path_valid(u'1' + ext) else u'.emb'

        fn = u'{0}{1}'.format(random.randint(1, 100), ext)
        return dump_sub_file(host_fullpath, fn, data_or_fileobj_to_write)

    if os.path.exists(sub_fullpath):
        os.remove(sub_fullpath)
    with open(sub_fullpath, 'wb') as fw:
        if hasattr(data_or_fileobj_to_write, u'read'):
            shutil.copyfileobj(fsrc=data_or_fileobj_to_write,
                               fdst=fw)
        else:
            fw.write(data_or_fileobj_to_write)
    return sub_fullpath


def extract_macros_from_office2003(fullpath, fileobj=None):
    '''

    :return: [(host_fullpath, filename_from_host, data), ... ]
    '''
    from oletools.olevba import VBA_Parser

    vp = VBA_Parser(fullpath, data=fileobj.read() if fileobj else None)

    r = []

    try:
        if vp.detect_vba_macros():
            macros = vp.extract_all_macros()
            assert (macros)  # macros detect, if cannot extact, must be error occured
            if macros:
                for (subfullpath, stream_path, vba_filename, vba_code) in macros:
                    a = os.path.basename(fullpath)
                    b = os.path.basename(subfullpath)
                    vba_filename += u'.vba'
                    sub = (
                        io_in_arg(fullpath),
                        io_in_arg(vba_filename if a == b else u'{0}_{1}'.format(b, vba_filename)),
                        vba_code
                    )
                    r.append(sub)
    except:
        pass
    finally:
        vp.close()

    return r


def extract_office2003_from_unknown_office(fullpath, fileobj=None):
    '''
    
    从不明 office(可能是 office2003, office2007) 中解出内嵌的 office2003
    :return: [
            (host_fullpath,filename_from_host,<file_open_handler>),
            ]
    '''

    import zipfile
    import olefile
    import io

    r = []

    if olefile.isOleFile(fileobj if fileobj else fullpath):
        r.append(
            (fullpath,
             os.path.basename(fullpath),
             fileobj if fileobj else  open(fullpath, 'rb'))
        )

    elif zipfile.is_zipfile(fileobj if fileobj else fullpath):
        with zipfile.ZipFile(fileobj if fileobj else fullpath) as z:
            for subfile in z.namelist():
                with z.open(subfile) as zt:
                    magic = zt.read(len(olefile.MAGIC))
                    if magic == olefile.MAGIC:
                        r.append(
                            (fullpath,
                             io_in_arg(subfile),
                             io.BytesIO(z.open(subfile).read()))
                        )
    else:
        raise ValueError(u'not office file')

    return r


def extract_subfile_in_10native_from_office2003(fullpath, fileobj=None):
    '''
    :param fileobj_or_fullpath: 
    :return: None / (file_name, file_content)
    '''
    import olefile
    from office_10native import escape_office_10native_from_olefileio
    f = olefile.OleFileIO()
    f.open(fileobj if fileobj else fullpath)
    r = escape_office_10native_from_olefileio(f)
    f.close()
    if r:
        return [(fullpath, r[0] + u'1', r[-1])]
    return None


def extract_subfile_in_10native_from_unknown_office(fullpath, fileobj=None):
    rs = extract_office2003_from_unknown_office(fullpath, fileobj)
    r = []

    _func = lambda e: extract_subfile_in_10native_from_office2003(
        fullpath=e[0] if os.path.basename(e[0]) == e[1] else  u'{0}_{1}'.format(e[0], io_path_format(e[1], u'_')),
        fileobj=e[2]
    )
    for ev in rs:
        r.extend(_func(ev))
    return r


def _extract_attachment_from_attachment(attachment, depth, results):
    '''
    call by others, and also call by self
    
    :return: 
    '''

    from base64_to_office import decode_mso_to_office, is_mso_buffer

    fn = attachment.get_filename()
    fn = io_in_arg(fn)
    if fn is None:
        v = attachment.get(u'Content-Location', None)
        if v:
            fn = os.path.split(v)[-1]
    if not fn:
        fn = u'noname.emb'
    fn = u'{0:0<3}.{1}'.format(depth, fn)

    if attachment.is_multipart():
        payloads = attachment.get_payload(decode=False)
        depth *= 10
        for e in payloads:
            depth += 1
            _extract_attachment_from_attachment(e, depth, results)
    else:
        data = attachment.get_payload(decode=True)

        if is_mso_buffer(data):
            ole = decode_mso_to_office(data)
            if ole:
                fn = fn + u'.office'
                results.append((fn, ole))
            else:
                results.append((fn, data))
        else:
            # results.append((fn, attachment.get_payload(decode=False)))
            results.append((fn, data))


def extract_attachment_from_eml(fullpath):
    '''
    
    :return: [(host_fullpath, filename_from_host, file_content)]
    '''
    import email
    r = []
    with open(fullpath, 'rb') as f:
        f_eml = email.message_from_file(f)
        attachments = f_eml.get_payload()
        if isinstance(attachments, list):
            depth = 1
            for attachment in attachments:
                _extract_attachment_from_attachment(attachment, depth, r)
                depth += 1
        else:
            _extract_attachment_from_attachment(f_eml, 1, r)

    return [(fullpath,) + e for e in r]


def extract_attachment_from_msg(fullpath):
    '''
    
    :return: [(host_fullpath, filename_from_host, file_content)]
    '''
    from ExtractMsg import Message
    msg = Message(fullpath)
    r = []
    for attachment in msg.attachments:
        name = attachment.longFilename
        # name = u'{0}_{1}'.format(fullpath, name)
        r.append((fullpath, io_in_arg(name), attachment.data))
    return r


def dump_framework(files, pfn_extract):
    c_ok = 0
    c_fail = 0

    for e in files:
        io_sys_stdout(e)
        io_sys_stdout(u'->')
        try:
            r = pfn_extract(e)
            map(lambda ev: dump_sub_file(*ev), r)
            c_ok += 1
            io_print('')
        except Exception as e:
            c_fail += 1
            io_print(u'fail,{}'.format(repr(e)))

    io_print(u'[+] ok {} fail {}'.format(c_ok, c_fail))


def unit_test():
    fullpath_sample = os.path.join(curpath, u'unit_test_sample')

    test_list = [
        # function , fullpath, subfiles_hash
        (extract_attachment_from_eml, os.path.join(fullpath_sample, u'eml'), set([u'470df372067d81e01159c2b1681ef9dc',
                                                                                  u'ff977f569030dfc524a2111e1b517a6b'])),
        (extract_macros_from_office2003, os.path.join(fullpath_sample, u'macros_from_office2003'),
         set([u'2633e74cb489ff3b9a606cb4885d6831'])),
        (extract_subfile_in_10native_from_office2003, os.path.join(fullpath_sample, u'emb_pe'),
         set([u'74d223c40b49a6cbb9494783a61aa707'])),
        (extract_subfile_in_10native_from_unknown_office, os.path.join(fullpath_sample, u'olex_emb_pe'),
         set([u'74d223c40b49a6cbb9494783a61aa707']))
    ]

    for e in test_list:
        fs = e[0](e[1])
        if fs:
            for fp in fs:
                fp = fp[-1]
                if hasattr(fp, u'read'):
                    h = io_hash_stream(fp)
                else:
                    h = io_hash_memory(fp)
                assert (h in e[2])

    io_print(u'pass unit test')


def entry():
    '''
    argparse.ArgumentParser 说明：
      add_argument('path', , help='file path') 叫增加位置参数，参数的值由传入顺序而定， 且不可以使用 --path 前缀做引导
      add_argument('--path', , help='file path') 叫增加可选参数，参数的值由前缀引导而定，要使其必选，需要使用 required=True 指定

    :return:
    '''
    parser = argparse.ArgumentParser()
    parser.add_argument('-e', '--embeddings', action='store_true', help='extract embeddings')
    parser.add_argument('-m', '--msg', action='store_true', help='extract attachment in msg')
    parser.add_argument('--eml', action='store_true', help='extract attachment in eml')
    parser.add_argument('restargs', nargs='+')
    args = parser.parse_args()

    fs = io_files_from_arg(args.restargs)
    if args.embeddings:
        dump_framework(fs, extract_subfile_in_10native_from_office2003)
    elif args.msg:
        dump_framework(fs, extract_attachment_from_msg)
    elif args.eml:
        dump_framework(fs, extract_attachment_from_eml)
    else:
        dump_framework(fs, extract_macros_from_office2003)


if __name__ == '__main__':
    # unit_test()
    entry()
