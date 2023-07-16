#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import paramiko
import getpass
import time
import datetime
import os
import sys
import threading
import logging
import 表格文件读取

if os.name == 'nt':
    import msvcrt
else:
    import tty
    import termios

成功数 = 0
错误数 = 0
命令错误数 = 0


# 命令超时错误捕获
class 命令超时错误(Exception):
    pass


class 自定义错误(Exception):
    pass


# 命令执行模块
def 执行命令(shell, 主机名, ssh日志文件, 命令们=None, 结束们=None, 命令超时时间=60):
    if 命令们 is None:
        命令们 = {'cmd1': 'dis ver', 'cmd2': 'dis th'}
    if 结束们 is None:
        结束们 = {}

    global 命令错误数

    输出 = ''

    # 执行自定义命令
    for 命令编号, 命令 in 命令们.items():
        中断计数 = 0
        命令 = 命令.replace(r'\n', '\n')
        结束 = 命令编号.replace(r'cmd', 'end')
        shell.send((命令 + '\n').encode('utf-8'))
        time.sleep(0.5)

        while True:
            if shell.recv_ready():
                输出 += shell.recv(1024).decode('utf-8')
                if 结束 in 结束们:
                    if 结束们[结束] in 输出:
                        break
                else:
                    if '<' + 主机名 in 输出:
                        break
                    elif '[' + 主机名 in 输出:
                        break
            if 中断计数 >= 命令超时时间:
                with open(ssh日志文件, 'a', encoding='utf-8') as f_sshlog:
                    f_sshlog.write(f'{输出}')
                命令错误数 += 1
                raise 命令超时错误(f'{命令}')
            中断计数 += 1
            time.sleep(1)
        with open(ssh日志文件, 'a', encoding='utf-8') as f_sshlog:
            f_sshlog.write(f'{输出}')
        输出 = ''

    return True


def ssh(主机名=None, 地址='127.0.0.1', 端口=22, 用户名='admin', 密码们=None, 超时时间=20,
        命令们=None, 结束们=None, 命令超时时间=60):
    if 密码们 is None:
        密码们 = ['admin@123']
    if 命令们 is None:
        命令们 = {'cmd1': 'dis ver', 'cmd2': 'dis th'}
    if 结束们 is None:
        结束们 = {}

    # 配置logging记录paramiko详细日志，不需要记录可以注释
    logging.basicConfig(level=logging.DEBUG, filename=f'log/paramiko-{datetime.date.today()}.log', filemode='w',
                        format='%(asctime)s - %(levelname)s - %(message)s')

    global 成功数, 错误数

    日志缓存 = []
    已连接 = False
    中断标记 = False

    # 创建日志目录和文件
    os.makedirs('log', exist_ok=True)
    os.makedirs('SSH_log', exist_ok=True)
    日志文件 = f'log/{datetime.date.today()}.txt'
    ssh日志文件 = f'SSH_log/{datetime.date.today()}-{地址}.txt'

    # 创建SSH客户端
    client = paramiko.SSHClient()

    # 设置密钥验证目录和文件
    # os.makedirs('SSH_key', exist_ok=True)
    # 已知主机文件 = './SSH_key/known_hosts'
    # with open(已知主机文件, 'w') as f:
    #     pass

    # 验证本地密钥
    # client.set_missing_host_key_policy(paramiko.WarningPolicy())
    # client.load_host_keys(os.path.expanduser(known_hosts文件))

    # 自动添加主机密钥
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    for 密码 in 密码们:
        try:
            print(f'{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]} 尝试连接{地址}')
            日志缓存.append(f'{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]} 尝试连接{地址} ')
            client.connect(hostname=地址, port=端口, username=用户名, password=密码, timeout=超时时间)

            已连接 = True
            print(f'{地址}连接成功')
            日志缓存.append(f'连接成功 ')

            # 创建交互式shell
            shell = client.invoke_shell()

            time.sleep(0.5)
            输出 = ''

            # 获取主机名
            if 主机名 is None:
                主机名 = 0
                while True:
                    if shell.recv_ready():
                        输出 += shell.recv(1024).decode('utf-8')

                        if '<' in 输出 and '>' in 输出:
                            主机名 = 输出.splitlines()[-1][1:-1]
                            print(主机名)
                            日志缓存.append(f'{主机名} ')
                            break

                    if 主机名 > 10:
                        break
                    主机名 += 1
                    time.sleep(1)

                if 主机名.isdigit():
                    print(f'获取设备名称错误跳过{地址}')
                    日志缓存.append(f'获取设备名称错误跳过{地址}\n')
                    break

            # 设置临时输出长度简化获取输出处理，仅在华为上测试过！
            shell.send(b'screen-length 0 temporary\n')  # 设置输出长度为0，临时模式
            time.sleep(1)
            输出 += shell.recv(1024).decode('utf-8')

            # 写入自定义命令执行前的ssh日志，不需要可以注释掉：
            with open(ssh日志文件, 'a', encoding='utf-8') as f_sshlog:
                f_sshlog.write(输出)

            # 执行自定义命令
            执行 = 执行命令(shell, 主机名, ssh日志文件, 命令们, 结束们, 命令超时时间)
            if 执行:
                print(f'执行完成')
                日志缓存.append(f'执行完成 ')
                成功数 += 1

        except paramiko.AuthenticationException:
            print(f'身份验证失败：{地址}')
            日志缓存.append(f'身份验证失败：{地址}\n')

        except paramiko.SSHException as e:
            print(f'SSH连接错误：{地址}, {str(e)}')
            日志缓存.append(f'SSH连接错误：{地址}, {str(e)}\n')
            中断标记 = True

        except paramiko.ssh_exception.NoValidConnectionsError as e:
            print(f'无法连接到主机：{地址}, {str(e)}')
            日志缓存.append(f'无法连接到主机：{地址}, {str(e)}\n')
            中断标记 = True

        except TimeoutError as e:
            print(f'连接超时：{地址}, {str(e)}')
            日志缓存.append(f'连接超时：{地址}, {str(e)}\n')
            中断标记 = True

        except 命令超时错误 as e:
            日志缓存.append(f'命令执行超时，跳过主机：{地址}, {str(e)}\n')

        except Exception as e:
            print(f'未知错误：{地址}, {str(e)}')
            日志缓存.append(f'遇到未知错误：{地址}, {str(e)}\n')

        finally:
            if not 已连接:
                with open(日志文件, 'a', encoding='utf-8') as f_log:
                    for 日志 in 日志缓存:
                        f_log.write(日志)
                日志缓存 = []
                if 中断标记:
                    break
                time.sleep(0.5)
            elif 已连接:
                client.close()
                print(f'关闭连接：{地址}')
                日志缓存.append(f'关闭连接：{地址}\n')
                with open(日志文件, 'a', encoding='utf-8') as f_log:
                    for 日志 in 日志缓存:
                        f_log.write(日志)
                break

    # 错误计数
    if not 已连接:
        错误数 += 1


def txt():
    # 创建日志目录和文件
    os.makedirs('log', exist_ok=True)
    日志文件 = f'log/{datetime.date.today()}.txt'

    print(f'程序启动时间：{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}')
    with open(日志文件, 'a', encoding='utf-8') as f_log:
        f_log.write(f'程序启动时间：{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}\n')

    用户名 = input('请输入用户名：')
    if 用户名 == '':
        用户名 = 'admin'

    # 密码 = input('请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    密码们 = getpass.getpass(prompt='请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    if 密码们 == ['']:
        密码们 = ['admin@123']
    密码们 = 密码们[:3]

    # 读取主机地址列表
    with open('host.txt', 'r') as f:
        主机们 = f.read().splitlines()

    # 读取命令
    命令编号 = 1
    命令们 = {}
    with open('commands.txt', 'r') as f:
        命令 = f.read().splitlines()
    for 命令 in 命令:
        命令们['cmd' + str(命令编号)] = 命令
        命令编号 += 1

    # 创建线程
    线程们 = []
    for 主机 in 主机们:
        线程 = threading.Thread(target=ssh, kwargs={'地址': 主机, '用户名': 用户名, '密码们': 密码们, '命令们': 命令们})
        线程们.append(线程)
        # time.sleep(0.2)
        线程.start()

    # 等待线程完成
    for 线程 in 线程们:
        线程.join()

    print(f'运行结束,共{str(len(主机们))}台主机，成功{成功数}台，错误{错误数}台，命令执行错误{命令错误数}台')
    with open(日志文件, 'a', encoding='utf-8') as f_log:
        f_log.write(f'运行结束,共{str(len(主机们))}台主机，成功{成功数}，错误{错误数}台，命令执行错误{命令错误数}台\n')


def excel(excel文件='SSH.xlsx', 表名='Sheet1'):
    try:
        主机们, 返回信息 = 表格文件读取.处理Excel文件(excel文件, 表名, True)
        for 信息 in 返回信息:
            print(信息)
        print(f'共{str(len(主机们))}台主机')
    except Exception as e:
        raise 自定义错误(e)

    # 创建日志目录和文件
    os.makedirs('log', exist_ok=True)
    日志文件 = f'log/{datetime.date.today()}.txt'

    # 密码 = input('请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    密码们 = getpass.getpass(prompt='请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    if 密码们 == ['']:
        密码们 = ['admin@123']
    密码们 = 密码们[:3]

    线程们 = []
    命令们 = {}
    结束们 = {}
    for 主机 in 主机们:
        for 命令编号, 命令 in 主机.items():
            if 'cmd' in 命令编号:
                命令们[命令编号] = 命令
            if 'end' in 命令编号:
                结束们[命令编号] = 命令
        线程 = threading.Thread(target=ssh,
                                kwargs={'主机名': 主机['Hostname'], '地址': 主机['Address'], '用户名': 主机['Username'],
                                        '密码们': 密码们, '命令们': 命令们, '结束们': 结束们})
        线程们.append(线程)
        # time.sleep(0.2)
        线程.start()

    for 线程 in 线程们:
        线程.join()

    print(f'运行结束,共{str(len(主机们))}台主机，成功{成功数}台，错误{错误数}台，命令执行错误{命令错误数}台')
    with open(日志文件, 'a', encoding='utf-8') as f_log:
        f_log.write(f'运行结束,共{str(len(主机们))}台主机，成功{成功数}，错误{错误数}台，命令执行错误{命令错误数}台\n')

    return True


# 清理屏幕
def 清屏():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


# 键盘阻塞
def 键盘阻塞():
    if os.name == 'nt':
        while True:
            if msvcrt.kbhit():
                msvcrt.getch()
                # 忽略任何键盘输入
                # 可以根据需要进行逻辑处理
                continue
            break


# 输入错误
def 键盘输入错误():
    清屏()
    键盘阻塞()
    print('输入有误请重新输入')
    time.sleep(1)


# 获取键盘输入
def 获取键盘输入():
    if os.name == 'nt':
        按键 = msvcrt.getch().decode('utf-8')
        return 按键
    else:
        文件描述符 = sys.stdin.fileno()
        旧设置 = termios.tcgetattr(文件描述符)
        try:
            tty.setraw(文件描述符)
            按键 = sys.stdin.read(1)
        finally:
            termios.tcsetattr(文件描述符, termios.TCSADRAIN, 旧设置)
        return 按键


# 选择文件读取模式
def 选择文件读取模式():
    excel文件路径 = 'SSH.xlsx'
    表名 = 'Sheet1'
    返回值 = False
    while True:
        清屏()
        print("请选择模式：")
        print("  1.txt模式")
        print("  2.excel模式")
        按键 = 获取键盘输入()

        if 按键 == '1':
            清屏()
            print('txt 模式')
            txt()
            break

        elif 按键 == '2':
            清屏()
            while True:
                清屏()
                print(f'excel模式：')
                print(f'  1.文件名：{excel文件路径}')
                print(f'  2.表名：{表名}')
                print(f'  3.继续')
                print(f'  0.返回上一层')
                按键 = 获取键盘输入()

                if 按键 == '1':
                    清屏()
                    新文件 = input(f'当前文件名 {excel文件路径}，请输入新的文件名：\n')
                    if 新文件:
                        excel文件路径 = 新文件
                elif 按键 == '2':
                    清屏()
                    新表名 = input(f'当前表名 {表名}，请输入新的表名：\n')
                    if 新表名:
                        表名 = 新表名
                elif 按键 == '3':
                    清屏()
                    try:
                        返回值 = excel(excel文件路径, 表名)
                    except Exception as e:
                        print('')
                        print(e)
                        print('')
                        print('excel文件错误，按任意键返回。按esc退出...')
                        print('')
                        按键 = 获取键盘输入()
                        if 按键 == chr(27):  # chr(27)是esc
                            清屏()
                            sys.exit(1)
                        else:
                            continue
                    break
                elif 按键 == '0':
                    break
                elif 按键 == chr(27):  # chr(27)是esc
                    break
                else:
                    键盘输入错误()

            if 返回值:
                break

        elif 按键 == chr(27):  # chr(27)是esc
            清屏()
            sys.exit()

        else:
            键盘输入错误()

    键盘阻塞()
    print('完成！按任意键退出...')
    获取键盘输入()


if __name__ == "__main__":
    选择文件读取模式()
