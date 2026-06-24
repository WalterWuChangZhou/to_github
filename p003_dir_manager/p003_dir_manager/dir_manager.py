from pathlib import Path
import csv
import shutil
from datetime import datetime
import inspect
from typing import List, Optional


class DirManager:
    """数据管理器 - 自动创建目录和CSV，支持备份和自动清理"""

    def __init__(self,
                 script_path=None,
                 headers: List[str] | None= None,
                 auto_backup: bool = True,
                 auto_cleanup: bool = True,
                 keep_backups: int = 5):
        """
        初始化数据管理器

        Args:
            script_path: 脚本路径，传入 __file__ 以使用调用脚本的名称
            headers: CSV表头，默认为 ['date', 'currency', 'rate']
            auto_backup: 是否在每次添加数据时自动备份
            auto_cleanup: 是否自动清理旧备份
            keep_backups: 保留的备份数量，默认保留5个
        """
        # 获取脚本路径
        if script_path is None:
            frame = inspect.currentframe().f_back
            caller_file = frame.f_globals.get('__file__', '')
            if caller_file:
                script_path = Path(caller_file).resolve()
            else:
                script_path = Path.cwd() / "unknown"
        else:
            script_path = Path(script_path).resolve()

        # 设置参数
        self.headers = headers if headers is not None else ['date', 'currency', 'rate']
        self.auto_backup = auto_backup
        self.auto_cleanup = auto_cleanup
        self.keep_backups = keep_backups

        # 设置路径
        self.script_dir = script_path.parent
        self.script_name = script_path.stem

        self.data_dir = self.script_dir / f"{self.script_name}_DATA"
        self.data_file = self.data_dir / f"{self.script_name}.csv"
        self.backup_dir = self.data_dir / f"{self.script_name}_BACKUP"

        # 创建目录
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.backup_dir.mkdir(parents=True, exist_ok=True)

        # 如果文件不存在，创建并写入表头
        if not self.data_file.exists():
            self._create_file()
        else:
            # 检查表头是否匹配
            self._check_headers()

        # 如果启用自动清理，在初始化时清理一次
        if self.auto_cleanup:
            print(f"🧹 初始化时清理旧备份 (保留 {self.keep_backups} 个)...")
            self.delete_old_backups(keep=self.keep_backups)

    def _create_file(self):
        """创建CSV文件并写入表头"""
        with open(self.data_file, 'w', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerow(self.headers)

    def _check_headers(self):
        """检查现有文件的表头是否匹配"""
        try:
            with open(self.data_file, 'r', encoding='utf-8-sig') as f:
                reader = csv.reader(f)
                existing_headers = next(reader, None)
                if existing_headers and existing_headers != self.headers:
                    print(f"⚠️ 警告: 现有表头 {existing_headers} 与自定义表头 {self.headers} 不匹配")
        except Exception as e:
            print(f"⚠️ 读取表头时出错: {e}")

    def add(self, *args, **kwargs):
        """
        添加一条数据

        用法1: add('2026-01-17', 'USD', 7.2345)  # 按位置传参
        用法2: add(date='2026-01-17', currency='USD', rate=7.2345)  # 按关键字传参
        用法3: add({'date': '2026-01-17', 'currency': 'USD', 'rate': 7.2345})  # 字典
        """
        # 如果启用自动备份，在写入前备份
        if self.auto_backup:
            self.backup()
            # 备份后自动清理
            if self.auto_cleanup:
                self.delete_old_backups(keep=self.keep_backups)

        # 处理各种输入方式
        if len(args) == 1 and isinstance(args[0], dict):
            # 字典方式
            row = [args[0].get(header, '') for header in self.headers]
        elif kwargs:
            # 关键字参数方式
            row = [kwargs.get(header, '') for header in self.headers]
        elif len(args) == len(self.headers):
            # 位置参数与表头数量匹配
            row = list(args)
        else:
            # 尝试按顺序匹配
            row = list(args)
            if len(row) > len(self.headers):
                print(f"⚠️ 警告: 数据列数 ({len(row)}) 超过表头列数 ({len(self.headers)})")
                row = row[:len(self.headers)]
            elif len(row) < len(self.headers):
                # 不足的列用空值填充
                row += [''] * (len(self.headers) - len(row))

        # 写入数据
        try:
            with open(self.data_file, 'a', encoding='utf-8-sig', newline='') as f:
                csv.writer(f).writerow(row[:len(self.headers)])
        except Exception as e:
            print(f"❌ 写入数据失败: {e}")

    def add_dict(self, data: dict):
        """
        使用字典添加数据

        Args:
            data: 字典，键必须与表头匹配
        """
        if self.auto_backup:
            self.backup()
            if self.auto_cleanup:
                self.delete_old_backups(keep=self.keep_backups)

        row = [data.get(header, '') for header in self.headers]
        with open(self.data_file, 'a', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerow(row)

    def add_list(self, row: list):
        """
        使用列表添加数据（按顺序）

        Args:
            row: 数据列表，按表头顺序排列
        """
        if self.auto_backup:
            self.backup()
            if self.auto_cleanup:
                self.delete_old_backups(keep=self.keep_backups)

        if len(row) < len(self.headers):
            row = list(row) + [''] * (len(self.headers) - len(row))
        row = row[:len(self.headers)]

        with open(self.data_file, 'a', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerow(row)

    def read(self):
        """读取所有数据"""
        try:
            with open(self.data_file, 'r', encoding='utf-8-sig') as f:
                return list(csv.reader(f))
        except Exception as e:
            print(f"❌ 读取数据失败: {e}")
            return []

    def get_headers(self):
        """获取当前表头"""
        return self.headers

    def backup(self):
        """备份数据"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_file = self.backup_dir / f"{self.script_name}_backup_{timestamp}.csv"
        try:
            shutil.copy2(self.data_file, backup_file)
            print(f"💾 已创建备份: {backup_file.name}")
            return backup_file
        except Exception as e:
            print(f"❌ 备份失败: {e}")
            return None

    def delete_old_backups(self, keep: int = 5, verbose: bool = True):
        """
        删除旧备份，保留最近的N个

        Args:
            keep: 保留的备份文件数量，默认保留5个
            verbose: 是否输出详细信息

        Returns:
            list: 被删除的文件路径列表
        """
        # 获取所有备份文件
        backup_files = list(self.backup_dir.glob(f"{self.script_name}_backup_*.csv"))

        if not backup_files:
            if verbose:
                print("📁 没有找到备份文件")
            return []

        # 按修改时间排序（旧的在前）
        backup_files.sort(key=lambda x: x.stat().st_mtime)

        if verbose:
            print(f"📁 找到 {len(backup_files)} 个备份文件，保留 {keep} 个")

        # 如果备份文件数量 <= 保留数量，不做任何操作
        if len(backup_files) <= keep:
            if verbose:
                print(f"✅ 备份文件数量 ({len(backup_files)}) 未超过保留数量 ({keep})，无需删除")
            return []

        # 确定要删除的旧文件（保留最新的 keep 个）
        files_to_delete = backup_files[:-keep]

        # 执行删除
        deleted_files = []
        for file_path in files_to_delete:
            try:
                file_path.unlink()
                deleted_files.append(str(file_path))
                if verbose:
                    print(f"  🗑️ 删除: {file_path.name}")
            except Exception as e:
                print(f"  ❌ 删除失败 {file_path.name}: {e}")

        if verbose and deleted_files:
            print(f"✅ 已清理 {len(deleted_files)} 个旧备份，保留 {keep} 个最新备份")

        return deleted_files

    def get_backup_count(self):
        """获取当前备份文件数量"""
        return len(list(self.backup_dir.glob(f"{self.script_name}_backup_*.csv")))

    def list_backups(self):
        """列出所有备份文件"""
        backup_files = sorted(
            self.backup_dir.glob(f"{self.script_name}_backup_*.csv"),
            key=lambda x: x.stat().st_mtime,
            reverse=True
        )

        if not backup_files:
            print("📁 没有找到备份文件")
            return []

        print(f"\n📋 备份文件列表 (共 {len(backup_files)} 个):")
        for i, file_path in enumerate(backup_files, 1):
            size = file_path.stat().st_size / 1024
            mtime = datetime.fromtimestamp(file_path.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S')
            print(f"  {i}. {file_path.name} ({size:.2f} KB) - {mtime}")

        return backup_files