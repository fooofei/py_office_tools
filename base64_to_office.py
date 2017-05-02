# coding=utf-8


def is_mso_buffer(file_content):
    return file_content and file_content.startswith("ActiveMime") or False


def io_decode_base64(base64_content):
    '''
            
    :param data_base64:
    :return: decoded bas64 data
    '''
    import binascii
    return binascii.a2b_base64(base64_content)


def decode_mso_to_office(mso_content):
    '''
    
    eml 类型的文件中 会遇到 base64 ， 解码之后是 ole 或者 olex
    另外的例子 https://github.com/phishme/python-amime/
    
    :return: None / office content
    '''
    import struct
    import zlib

    if not is_mso_buffer(mso_content):
        return None

    try:
        offset = struct.unpack_from('<H', mso_content, offset=0x1E)[0] + 46
        for start in (offset, 0x32, 0x22A):
            try:
                return zlib.decompress(mso_content[start:])
            except:
                return None
    except:
        return None
