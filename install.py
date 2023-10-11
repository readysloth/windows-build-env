import re
import sys
import abc
import shutil
import asyncio
import logging
import zipfile
import threading
import subprocess as sp

from pathlib import Path
from urllib.request import urlretrieve

ENV_REGISTRY_KEY = r'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment'

DEV_TOOLS_PATH = Path(r'C:\dev_tools')

# Не больше трех подключений, серверу плохо
DOWNLOAD_SEMAPHORE = threading.Semaphore(3)

LOGGER = logging.getLogger(__name__)
handler = logging.StreamHandler(sys.stdout)
handler.setFormatter(logging.Formatter('%(asctime)s: %(message)s'))
LOGGER.setLevel(logging.INFO)
LOGGER.addHandler(handler)


def registry_create(path, key, value, node_type):
    append_args = ['reg', 'add', path, '/v', key, '/t', node_type,
                   '/d', value, '/f']
    sp.run(append_args)
    LOGGER.info(f'Registry: {value} created in {path}[{key}]')


def registry_append(path, key, value, delim=';'):
    def query_registry():
        query_args = ['reg', 'query', path, '/v', key]
        proc = sp.run(query_args, stdout=sp.PIPE)
        previous_value = [line for line in proc.stdout.decode().split('\r\n') if line][1]
        return [s for s in re.split(r'\s{2,}', previous_value) if s]

    var_name, var_type, var_value = query_registry()
    new_value = var_value + delim + value
    registry_create(path, key, new_value, var_type)
    return query_registry()


def generic_install(file, install_opts=None, install_append=None, install_post_append=None):
    install_args = ['cmd', '/c', 'start', '/wait']
    if install_append:
        install_args += install_append
    subprocess_args = install_args \
                      + [file] \
                      + (install_opts or []) \
                      + (install_post_append or [])
    sp.run(subprocess_args)


def msi_install(file, install_opts=None):
    generic_install(file,
                    install_opts=install_opts,
                    install_append=['msiexec', '/i'],
                    install_post_append=['/qn'])


def download_file(url, *args, depth=10, **kwargs):
    with DOWNLOAD_SEMAPHORE:
        try:
            filename = urlretrieve(url, *args, **kwargs)
            return filename
        except Exception as e:
            if depth > 0:
                LOGGER.warn(f'Download from {url} failed with {e}. Retrying {depth}')
                return download_file(url, *args, depth=depth - 1, **kwargs)
            raise


class Package(abc.ABC):
    def __init__(self, *args, **kwargs):
        self.logger = logging.getLogger(type(self).__name__)
        handler = logging.StreamHandler(sys.stdout)
        handler.setFormatter(logging.Formatter('%(asctime)s: %(name)s: %(message)s'))
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(handler)

    def install(self, file):
        pass

    def download(self):
        return download_file(self.url, filename=self.filename)[0]

    async def ainstall(self):
        loop = asyncio.get_running_loop()
        self.logger.info('Starting download')
        try:
            file = await loop.run_in_executor(None, self.download)
        except Exception as e:
            self.logger.error(f'Failed to download. Installation aborted due to {e}')
            return
        self.logger.info('Download finished, starting installation')
        self.install(file)
        self.logger.info('installation finished')
        try:
            await loop.run_in_executor(None, Path(file).unlink)
            self.logger.error(f'Successfully removed {file}')
        except Exception as e:
            self.logger.error(f'Failed to remove {file} due to {e}')


class Exe(Package):
    def install(self, file, install_opts=None):
        generic_install(file, install_opts=install_opts)


class Msi(Package):
    def install(self, file, install_opts=None):
        msi_install(file, install_opts=install_opts)


class Zip(Package):
    def install(self, file, folder):
        Path(folder).mkdir(exist_ok=True)
        with zipfile.ZipFile(file, 'r') as zip:
            zip.extractall(folder)

    async def ainstall(self):
        loop = asyncio.get_running_loop()
        self.logger.info('Starting download')
        try:
            file = await loop.run_in_executor(None, self.download)
        except Exception as e:
            self.logger.error(f'Failed to download. Installation aborted due to {e}')
            return
        self.logger.info('Download finished, starting unpacking')
        await loop.run_in_executor(None, self.install, file)
        self.logger.info('installation finished')
        try:
            await loop.run_in_executor(None, Path(file).unlink)
            self.logger.error(f'Successfully removed {file}')
        except Exception as e:
            self.logger.error(f'Failed to remove {file} due to {e}')


class Move(Package):
    def install(self, file, folder, alias=None):
        folder_path = Path(folder)
        folder_path.mkdir(exist_ok=True)
        if alias:
            shutil.move(file, str(Path(folder_path) / Path(alias)))
            return
        shutil.move(file, str(folder_path))


class LLVM(Exe):
    ver = '16.0.0'
    exe = f'LLVM-{ver}-win64.exe'
    url = f'https://github.com/llvm/llvm-project/releases/download/llvmorg-{ver}/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['/S'])
        registry_append(ENV_REGISTRY_KEY, 'Path', r'C:\Program Files\LLVM\bin')


class Python(Exe):
    ver = '3.10.10'
    exe = f'python-{ver}-amd64.exe'
    url = f'https://www.python.org/ftp/python/{ver}/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['/quiet',
                               'InstallAllUsers=1',
                               'CompileAll=1',
                               'PrependPath=1',
                               'Include_launcher=1'])


class Git(Exe):
    ver = '2.40.0'
    exe = f'Git-{ver}-64-bit.exe'
    url = f'https://github.com/git-for-windows/git/releases/download/v{ver}.windows.1/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['/VERYSILENT',
                               '/NORESTART',
                               '/NOCANCEL',
                               '/SP-',
                               '/ALLUSERS',
                               '/CLOSEAPPLICATIONS',
                               '/RESTARTAPPLICATIONS',
                               '/COMPONENTS=ext\\shellhere,assoc,assoc_sh'])


class Perl(Msi):
    ver = '5.32.1.1'
    msi = f'strawberry-perl-{ver}-64bit.msi'
    url = f'https://strawberryperl.com/download/{ver}/{msi}'
    filename = msi

    def install(self, file):
        super().install(file)


class Cmake(Msi):
    ver = '3.26.2'
    msi = f'cmake-{ver}-windows-x86_64.msi'
    url = f'https://github.com/Kitware/CMake/releases/download/v{ver}/{msi}'
    filename = msi

    def install(self, file):
        super().install(file, ['ADD_CMAKE_TO_PATH=System'])


class Nsis(Exe):
    ver = '3.08'
    exe = f'nsis-{ver}-setup.exe'
    url = f'https://prdownloads.sourceforge.net/nsis/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['/S', '/NCRC'])


class Wix(Zip):
    ver = '311'
    zip = f'wix{ver}-binaries.zip'
    url = f'https://github.com/wixtoolset/wix3/releases/download/wix3112rtm/{zip}'
    filename = zip

    def install(self, file):
        install_path = r'C:\WIX'
        super().install(file, install_path)
        registry_create(ENV_REGISTRY_KEY, 'WixInstallPath', install_path, 'REG_EXPAND_SZ')
        registry_create(ENV_REGISTRY_KEY, 'WixTargetsPath', install_path, 'REG_EXPAND_SZ')


class Dependencies(Zip):
    ver = '1.11.1'
    zip = 'Dependencies_x64_Release.zip'
    url = f'https://github.com/lucasg/Dependencies/releases/download/v{ver}/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class DependencyWalker(Zip):
    zip = 'depends22_x64.zip'
    url = f'http://www.dependencywalker.com/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class Ninja(Zip):
    ver = '1.11.1'
    zip = 'ninja-win.zip'
    url = f'https://github.com/ninja-build/ninja/releases/download/{ver}/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class W64Devkit(Zip):
    ver = '1.19.0'
    zip = f'w64devkit-{ver}.zip'
    url = f'https://github.com/skeeto/w64devkit/releases/download/{ver}/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class DnSpy(Zip):
    ver = '6.3.0'
    zip = 'dnSpy-net-win64.zip'
    url = f'https://github.com/dnSpyEx/dnSpy/releases/download/v{ver}/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class X64Dbg(Zip):
    ver = '2023-06-30_14-26'
    zip = f'snapshot_{ver}.zip'
    url = f'https://github.com/x64dbg/x64dbg/releases/download/snapshot/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))
        registry_append(ENV_REGISTRY_KEY, 'Path', str(DEV_TOOLS_PATH / Path('release')))


class Msys2(Exe):
    date = '2023-07-18'
    ver = '20230718'
    exe = f'msys2-x86_64-{ver}.exe'
    url = f'https://github.com/msys2/msys2-installer/releases/download/{date}/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['install',
                               '--root',
                               'C:\\MSYS2',
                               '--confirm-command'])
        registry_append(ENV_REGISTRY_KEY, 'Path', r'C:\MSYS2\usr\bin')


class ConEmu(Exe):
    ver = '230724'
    ver1 = '23.07.24'
    exe = f'ConEmuSetup.{ver}.exe'
    url = f'https://github.com/Maximus5/ConEmu/releases/download/v{ver1}/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['/p:x64,adm', '/qn'])


class Putty(Msi):
    ver = '0.79'
    msi = f'putty-64bit-{ver}-installer.msi'
    url = f'https://the.earth.li/~sgtatham/putty/latest/w64/{msi}'
    filename = msi

    def install(self, file):
        super().install(file)


class Mitmproxy(Exe):
    ver = '10.1.1'
    exe = f'mitmproxy-{ver}-windows-x64-installer.exe'
    url = f'https://downloads.mitmproxy.org/{ver}/{exe}'
    filename = exe

    def install(self, file):
        super().install(file, ['--mode', 'unattended'])


class SnRemove(Zip):
    zip = 'snremove.zip'
    url = f'https://www.nirsoft.net/dot_net_tools/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class SysInternals(Zip):
    zip = 'SysinternalsSuite.zip'
    url = f'https://download.sysinternals.com/files/{zip}'
    filename = zip

    def install(self, file):
        super().install(file, str(DEV_TOOLS_PATH))


class EWDK(Zip):
    ver = '10'
    zip = 'EWDK11.zip'
    files = [f'x{letter}' for letter in 'abc']
    urls = [f'https://github.com/readysloth/msvc-wine/releases/download/v0.1.0/{file}'
            for file in files]
    filename = zip

    def download(self):
        for url in self.urls:
            self.url = url
            super().download()
        sp.run(['cmd', '/c', 'type'] + self.files + ['>', self.zip])
        for file in self.files:
            Path(file).unlink()


async def install_all(packages):
    install_coros = [p.ainstall() for p in packages]
    await asyncio.gather(*install_coros)


if __name__ == '__main__':
    PACKAGES = [LLVM(),
                Python(),
                Git(),
                Perl(),
                Cmake(),
                Nsis(),
                Wix(),
                Dependencies(),
                DependencyWalker(),
                Ninja(),
                W64Devkit(),
                DnSpy(),
                X64Dbg(),
                Msys2(),
                Putty(),
                Mitmproxy(),
                ConEmu(),
                SnRemove(),
                SysInternals(),
                EWDK(),
                ]
    asyncio.run(install_all(PACKAGES))
    registry_append(ENV_REGISTRY_KEY, 'Path', str(DEV_TOOLS_PATH))
    registry_create(ENV_REGISTRY_KEY, 'UseEnv', 'true', 'REG_EXPAND_SZ')
