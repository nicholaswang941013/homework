import sys

# 禁用Python緩存文件生成
sys.dont_write_bytecode = True

from registration.registration import show_registration_form

__all__ = ['show_registration_form'] 