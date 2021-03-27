def is_chinese(uchar):
    """判断一个unicode是否是汉字"""
    if uchar >= u'\u4e00' and uchar <= u'\u9fa5':
        return True
    else:
        return False

def strip_chinese(str_input):
    """只保留输入字符串中的汉字"""
    str_input = str(str_input)
    str_result = ''.join([char_ for char_ in str_input if is_chinese(char_)])
    return str_result


if __name__ == '__main__':
    str_test = "IV.四季度 "
    list_result = [is_chinese(char_)  for char_  in str_test]
    tuple_result = zip(str_test, list_result)
    print([tuple_ for tuple_ in tuple_result])
    print(strip_chinese(str_test))