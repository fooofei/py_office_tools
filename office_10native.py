# coding=utf-8

import struct
import olefile
from  io_in_out import *


def get_office_10native_stream(olefileio_obj):
    '''
    
    :param olefileio_obj: olefile.OleFileIO()
    :return:  file buffer / None
    '''
    _fn = lambda iof, n: iof.openstream(n).read() if iof.exists(n) else None

    r = _fn(olefileio_obj, u'\x01Ole10Native')
    if not r:
        ns = [u'ObjectPool', u'_1525708454', u'\x01Ole10Native']
        return _fn(olefileio_obj, ns)
    return r


def escape_office_10native_from_buffer(stream_buffer):
    '''
    :return: None / ('','','','') 
     
    解出 ole 中的 pe 文件
    ref https://raw.githubusercontent.com/unixfreak0037/officeparser/master/officeparser.py
    上面的有错误， 利用下面微软的文章修正
    ref https://code.msdn.microsoft.com/office/CSOfficeDocumentFileExtract-e5afce86
    '''
    size = struct.unpack('<L', stream_buffer[0:4])[0]
    data = stream_buffer[4:]

    unknown_short = None
    filename = []
    src_path = []
    dst_path = []
    actual_size = None
    unknown_long_1 = None
    unknown_long_2 = None
    # I thought this might be an OLE type specifier ???
    unknown_short = struct.unpack('<H', data[0:2])[0]
    data = data[2:]

    # filename
    i = 0
    while i < len(data):
        if ord(data[i]) == 0:
            break
        filename.append(data[i])
        i += 1
    filename = ''.join(filename)
    data = data[i + 1:]

    # source path
    i = 0
    while i < len(data):
        if ord(data[i]) == 0:
            break
        src_path.append(data[i])
        i += 1
    src_path = ''.join(src_path)
    data = data[i + 1:]

    # TODO I bet these next 8 bytes are a timestamp
    unknown_long_1 = struct.unpack('<L', data[0:4])[0]
    data = data[4:]

    # Next four bytes gives the size of the temporary path of the embedded file  in little endian format
    # This should be converted
    temp_path_size = struct.unpack('<L', data[0:4])[0]
    data = data[4:]

    # destination path? (interesting that it has my name in there)
    i = 0
    while i < len(data):
        if ord(data[i]) == 0:
            break
        dst_path.append(data[i])
        i += 1
    dst_path = ''.join(dst_path)

    # 修正第一个 ref 文章的 bug
    if len(dst_path) > temp_path_size:
        raise ValueError(u'stream decode error, len(dst_path)>temp_path_size ')

    data = data[temp_path_size:]

    # size of the rest of the data
    actual_size = struct.unpack('<L', data[0:4])[0]
    if not actual_size:
        return None
    data = data[4:]

    # (filename, <fullpath before put in ole>,<fullpath to write from ole>,data)
    filename = io_in_arg(filename)
    fullpath_original = io_in_arg(src_path)
    fullpath_dst = io_in_arg(dst_path)
    return (filename, fullpath_original, fullpath_dst, data[0:actual_size])


def escape_office_10native_from_olefileio(olefileio_obj):
    '''
    
    :return: same with escape_office_10native_from_buffer()
    '''
    r = get_office_10native_stream(olefileio_obj)
    return escape_office_10native_from_buffer(r) if r else None
