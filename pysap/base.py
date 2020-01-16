# -*- coding: utf-8 -*-
#  author: andrew budzan
# project: vf-corp-automation
#  script: base.py
#    date: 16.01.2020

from win32com.client import GetObject


class ComObject:
    def __init__(self, com):
        self._com = com
        self.sap_id = com.Id
        self.sap_name = com.Name
        self.sap_type = com.Type

    def __str__(self):
        return f'{self.sap_type} object:\n\tname: {self.sap_name}\n\tid: {self.sap_id}'

    def __repr__(self):
        return f'Sap{self.sap_type}'

    def __call__(self, *args, **kwargs):
        return self._com


class SapContainer(ComObject):
    def __init__(self, com):
        assert com.ContainerType is True, 'Object is not SapContainer'
        super().__init__(com)

    @property
    def children(self):
        return {i: ch for i, ch in enumerate(list(self().Children))}


class Sap(object):

    _pr_name = 'SAPGUI'

    def __new__(cls):
        if not hasattr(cls, '_instance'):
            cls._instance = super(Sap, cls).__new__(cls)
            engine = cls._get_engine()
            if engine:
                cls._instance.app = engine.GetScriptingEngine
            else:
                raise Exception(f'{cls._pr_name} is not launched. Log in and try again.')
        return cls._instance

    @staticmethod
    def _get_engine():
        try:
            engine = GetObject(Sap._pr_name)
        except:
            engine = None
        return engine



if __name__ == '__main__':
    s1 = Sap()
    ap = ComObject(s1.app)
    # print(ap().ContainerType)
    cont = SapContainer(s1.app)
    print(cont)
    print(cont.children[0].Id)
