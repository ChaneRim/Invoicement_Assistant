block_cipher = None

a = Analysis(['main.py'],
             pathex=['.'],
             binaries=[],
             datas=[('core/tax_core.py', 'core'),
                    ('view/main_ui.py', 'view'),
                    ('view/custom_info_del_ui.py', 'view'),
                    ('view/custom_info_change_ui.py', 'view'),
                    ('view/custom_info_add_ui.py', 'view'),
                    ('controller/main_controller.py', 'controller'),
                    ('controller/custom_info_del_controller.py', 'controller'),
                    ('controller/custom_info_change_controller.py', 'controller'),
                    ('controller/custom_info_add_controller.py', 'controller'),
                    ('file/config.json', 'file'),
                    ('file/pattern.json', 'file'),
                    ('data_center.py', 'data_center')],
             hiddenimports=[],
             hookspath=[],
             hooksconfig={},
             runtime_hooks=[],
             excludes=[],
             win_no_prefer_redirects=False,
             win_private_assemblies=False,
             cipher=block_cipher,
             noarchive=False)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          [],
          exclude_binaries=True,
          name='main',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          console=False)

coll = COLLECT(exe,
               a.binaries,
               a.zipfiles,
               a.datas,
               strip=False,
               upx=True,
               upx_exclude=[],
               name='main')