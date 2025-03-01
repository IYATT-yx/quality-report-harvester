from dialog import Dialog as dlg

import jieba
from jieba import posseg

class WordSegmentation:
    customDict = None

    @staticmethod
    def loadCustomDict(file: str) -> bool:
        """
        加载自定义词典

        Params:
            file: 自定义词典文件路径

        Returns:
            bool: 加载成功与否
        """
        if WordSegmentation.customDict is not None:
            dlg.log('删除自定义词典')
            for word in WordSegmentation.customDict:
                jieba.del_word(word)
                dlg.log(f'删除 {word}')

        WordSegmentation.customDict = []
        try:
            with open(file, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if len(line) == 0:
                        continue

                    parts = line.split()
                    match len(parts):
                        case 1:
                            tag = 'nr'
                        case 2:
                            tag = parts[1]
                        case _:
                            dlg.log(f'自定义词典格式错误（默认取第一个字段作为人名）: {line}', dlg.WARNING)
                            tag = 'nr'
                    word = parts[0]
                    if word not in WordSegmentation.customDict:
                        jieba.add_word(word, tag=tag)
                        WordSegmentation.customDict.append(word)
                    else:
                        dlg.log(f'自定义词典重复词: {line}', dlg.WARNING)
                        continue
        except FileNotFoundError:
            dlg.log(f'自定义词典文件不存在或无读取权限: {file}')
            return False
        except Exception as e:
            WordSegmentation.customDict = None
            dlg.log(f'加载自定义词典失败: {e}', dlg.ERROR)
            return False
        
        wordAndTags = WordSegmentation.queryCustomDict()
        if wordAndTags is not None:
            dlg.log('已加载自定义词典：')
            for word, tag in wordAndTags:
                dlg.log(f'{word} {tag}')
        else:
            dlg.log('自定义词典为空')
        
        return True

    @staticmethod
    def queryCustomDict() -> list[tuple[str, str]]:
        """
        查询已加载的自定义词典

        Returns:
            成功返回词和标签列表，空词典返回 None
        """
        if WordSegmentation.customDict is None:
            return None

        wordAndTags = []
        for word in wordAndTags:
            wordAndTag = posseg.cut(word)
            for word, tag in wordAndTag:
                wordAndTags.append((word, tag))
        return wordAndTags
    
    @staticmethod
    def do(text: str) -> list[str]:
        """
        分词处理，提取受处罚的人名和供应商

        Params:
            text: 待分词文本

        Returns:
            结果
        """
        names = []
        wordAndTags = posseg.cut(text)
        for word, tag in wordAndTags:
            if tag == 'nr':
                if '奖励' not in text:
                    names.append(word)
            elif tag == 'nt':
                names.append(word)
        return names
