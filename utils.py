# !/usr/bin/env python
"""
General Utilities
"""
import logging
import os
import sys
import configparser


__all__ = [
    "logging", "logger",
    "g_cur_path",
    "change_path_to_word_style",
    "g_cfg", 'Config'
]


g_cur_path = os.path.dirname(os.path.realpath(sys.argv[0]))
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)


def change_path_to_word_style(path):
    return path.replace('/', '\\')


class Config(object):
    ATTACHMENTS = 'attachments'
    GENERAL = 'general'
    SAVE_TEMP = 'save_temp'

    def __init__(self):
        self._cfg_path = cfg_path = os.path.join(g_cur_path, 'res\\setting.conf')
        self._cfg_parser = cfg_parser = configparser.ConfigParser()
        cfg_parser.read(cfg_path)
        self._init_cfg()

    def _init_cfg(self):
        cfg_parser = self._cfg_parser
        if self.ATTACHMENTS not in cfg_parser.sections():
            cfg_parser.add_section(self.ATTACHMENTS)
            cfg_parser.set(self.ATTACHMENTS, 't1', 'doc1;doc2;doc3')
        if self.GENERAL not in cfg_parser.sections():
            cfg_parser.add_section(self.GENERAL)
            cfg_parser.set(self.GENERAL, self.SAVE_TEMP, 'False')
        self.save()

    def save(self):
        with open(self._cfg_path, 'w') as f:
            self._cfg_parser.write(f)

    def set(self, section, option, value):
        self._cfg_parser.set(section, option, value)
        self.save()

    def add_section(self, section):
        self._cfg_parser.add_section(section)
        self.save()

    # 组合优于继承
    def __getattr__(self, item):
        return getattr(self._cfg_parser, item)


g_cfg = Config()

# g_cfg.options('attachments')
