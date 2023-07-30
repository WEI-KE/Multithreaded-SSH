#!/usr/bin/python3
# -*- coding: UTF-8 -*-

import datetime
import getpass
import logging
import os
import sys
import threading
import time
import paramiko
import excel_read_zh

if os.name == 'nt':
    import msvcrt
else:
    import tty
    import termios

success_count = 0
error_count = 0
command_error = 0


# 命令超时错误捕获
class CommandTimeoutError(Exception):
    pass


class CustomError(Exception):
    pass


# 命令执行模块
def execute_command(shell, hostname, ssh_log_file, commands=None, ends=None, command_timeout=60):
    if commands is None:
        commands = {'cmd1': 'dis ver', 'cmd2': 'dis th'}
    if ends is None:
        ends = {}

    global command_error

    output = ''

    # 执行自定义命令
    for command_num, command in commands.items():
        break_count = 0
        command = command.replace(r'\n', '\n')
        end = command_num.replace(r'cmd', 'end')
        shell.send((command + '\n').encode('utf-8'))
        time.sleep(0.5)

        while True:
            if shell.recv_ready():
                output += shell.recv(1024).decode('utf-8')
                if end in ends:
                    if ends[end] in output:
                        break
                else:
                    if '<' + hostname in output:
                        break
                    elif '[' + hostname in output:
                        break
            if break_count >= command_timeout:
                with open(ssh_log_file, 'a', encoding='utf-8') as f_sshlog:
                    f_sshlog.write(f'{output}')
                command_error += 1
                raise CommandTimeoutError(f'{command}')
            break_count += 1
            time.sleep(1)
        with open(ssh_log_file, 'a', encoding='utf-8') as f_sshlog:
            f_sshlog.write(f'{output}')
        output = ''

    return True


def ssh(hostname=None, address='127.0.0.1', port=22, username='admin', passwords=None, timeout=20,
        commands=None, ends=None, command_timeout=60):
    if passwords is None:
        passwords = ['admin@123']
    if commands is None:
        commands = {'cmd1': 'dis ver', 'cmd2': 'dis th'}
    if ends is None:
        ends = {}

    # 配置logging记录paramiko详细日志，不需要记录可以注释
    logging.basicConfig(level=logging.DEBUG, filename=f'log/paramiko-{datetime.date.today()}.log', filemode='w',
                        format='%(asctime)s - %(levelname)s - %(message)s')

    global success_count, error_count

    log_cache = []
    connected = False
    break_tag = False

    # 创建日志目录和文件
    os.makedirs('log', exist_ok=True)
    os.makedirs('SSH_log', exist_ok=True)
    log_file = f'log/{datetime.date.today()}.txt'
    ssh_log_file = f'SSH_log/{datetime.date.today()}-{address}.txt'

    # 创建SSH客户端
    client = paramiko.SSHClient()

    # 设置密钥验证目录和文件
    # os.makedirs('SSH_key', exist_ok=True)
    # known_hosts_file = './SSH_key/known_hosts'
    # with open(known_hosts_file, 'w') as f:
    #     pass

    # 验证本地密钥
    # client.set_missing_host_key_policy(paramiko.WarningPolicy())
    # client.load_host_keys(os.path.expanduser(known_hosts_file))

    # 自动添加主机密钥
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    for password in passwords:
        try:
            print(f'{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]} 尝试连接{address}')
            log_cache.append(f'{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]} 尝试连接{address} ')
            client.connect(hostname=address, port=port, username=username, password=password, timeout=timeout)

            connected = True
            print(f'{address}连接成功')
            log_cache.append(f'连接成功 ')

            # 创建交互式shell
            shell = client.invoke_shell()

            time.sleep(0.5)
            output = ''

            # 获取主机名
            if hostname is None:
                hostname = 0
                while True:
                    if shell.recv_ready():
                        output += shell.recv(1024).decode('utf-8')

                        if '<' in output and '>' in output:
                            hostname = output.splitlines()[-1][1:-1]
                            print(hostname)
                            log_cache.append(f'{hostname} ')
                            break

                    if hostname > 10:
                        break
                    hostname += 1
                    time.sleep(1)

                if hostname.isdigit():
                    print(f'获取设备名称错误跳过{address}')
                    log_cache.append(f'获取设备名称错误跳过{address}\n')
                    break

            # 设置临时输出长度简化获取输出处理，仅在华为上测试过！
            shell.send(b'screen-length 0 temporary\n')  # 设置输出长度为0，临时模式
            time.sleep(1)
            output += shell.recv(1024).decode('utf-8')

            # 写入自定义命令执行前的ssh日志，不需要可以注释掉：
            with open(ssh_log_file, 'a', encoding='utf-8') as f_sshlog:
                f_sshlog.write(output)

            # 执行自定义命令
            execute = execute_command(shell, hostname, ssh_log_file, commands, ends, command_timeout)
            if execute:
                print(f'执行完成')
                log_cache.append(f'执行完成 ')
                success_count += 1

        except paramiko.AuthenticationException:
            print(f'身份验证失败：{address}')
            log_cache.append(f'身份验证失败：{address}\n')

        except paramiko.SSHException as e:
            print(f'SSH连接错误：{address}, {str(e)}')
            log_cache.append(f'SSH连接错误：{address}, {str(e)}\n')
            break_tag = True

        except paramiko.ssh_exception.NoValidConnectionsError as e:
            print(f'无法连接到主机：{address}, {str(e)}')
            log_cache.append(f'无法连接到主机：{address}, {str(e)}\n')
            break_tag = True

        except TimeoutError as e:
            print(f'连接超时：{address}, {str(e)}')
            log_cache.append(f'连接超时：{address}, {str(e)}\n')
            break_tag = True

        except CommandTimeoutError as e:
            log_cache.append(f'命令执行超时，跳过主机：{address}, {str(e)}\n')

        except Exception as e:
            print(f'未知错误：{address}, {str(e)}')
            log_cache.append(f'遇到未知错误：{address}, {str(e)}\n')

        finally:
            if not connected:
                with open(log_file, 'a', encoding='utf-8') as f_log:
                    for log in log_cache:
                        f_log.write(log)
                log_cache = []
                if break_tag:
                    break
                time.sleep(0.5)
            elif connected:
                client.close()
                print(f'关闭连接：{address}')
                log_cache.append(f'关闭连接：{address}\n')
                with open(log_file, 'a', encoding='utf-8') as f_log:
                    for log in log_cache:
                        f_log.write(log)
                break

    # 错误计数
    if not connected:
        error_count += 1


def txt():
    # 创建日志目录和文件
    os.makedirs('log', exist_ok=True)
    log_file = f'log/{datetime.date.today()}.txt'

    print(f'程序启动时间：{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(f'程序启动时间：{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}\n')

    username = input('请输入用户名：')
    if username == '':
        username = 'admin'

    # password = input('请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    passwords = getpass.getpass(prompt='请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    if passwords == ['']:
        passwords = ['admin@123']
    passwords = passwords[:3]

    # 读取主机地址列表
    with open('host.txt', 'r') as f:
        hosts = f.read().splitlines()

    # 读取命令
    command_num = 1
    commands = {}
    with open('commands.txt', 'r') as f:
        cmds = f.read().splitlines()
    for cmd in cmds:
        commands['cmd' + str(command_num)] = cmd
        command_num += 1

    # 创建线程
    threads = []
    for address in hosts:
        thread = threading.Thread(target=ssh, kwargs={'address': address, 'username': username, 'passwords': passwords,
                                                      'commands': commands})
        threads.append(thread)
        # time.sleep(0.2)
        thread.start()

    # 等待线程完成
    for thread in threads:
        thread.join()

    print(f'运行结束,共{str(len(hosts))}台主机，成功{success_count}台，错误{error_count}台，命令执行错误{command_error}台')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(
            f'运行结束,共{str(len(hosts))}台主机，成功{success_count}，错误{error_count}台，命令执行错误{command_error}台\n')


def excel(excel_file='SSH.xlsx', sheet='Sheet1'):
    try:
        hosts, return_info = excel_read_zh.process_excel_file(excel_file, sheet, True)
        for info in return_info:
            print(info)
        print(f'{str(len(hosts))} hosts')
    except Exception as e:
        raise CustomError(e)

    # 创建日志目录和文件
    os.makedirs('log', exist_ok=True)
    log_file = f'log/{datetime.date.today()}.txt'

    print(f'程序启动时间：{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(f'程序启动时间：{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}\n')

    # password = input('请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    passwords = getpass.getpass(prompt='请输入设备密码用"空格"分隔（最多不超过3个）:').split(' ')
    if passwords == ['']:
        passwords = ['admin@123']
    passwords = passwords[:3]

    threads = []
    commands = {}
    ends = {}
    for host in hosts:
        for command_num, command in host.items():
            if 'cmd' in command_num:
                commands[command_num] = command
            if 'end' in command_num:
                ends[command_num] = command
        if host['Hostname']:
            thread = threading.Thread(target=ssh,
                                      kwargs={'hostname': host['Hostname'], 'address': host['Address'],
                                              'username': host['Username'], 'passwords': passwords,
                                              'commands': commands, 'ends': ends})
        else:
            thread = threading.Thread(target=ssh,
                                      kwargs={'address': host['Address'],
                                              'username': host['Username'], 'passwords': passwords,
                                              'commands': commands, 'ends': ends})
        threads.append(thread)
        # time.sleep(0.2)
        thread.start()

    for thread in threads:
        thread.join()

    print(f'运行结束,共{str(len(hosts))}台主机，成功{success_count}台，错误{error_count}台，命令执行错误{command_error}台')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(
            f'运行结束,共{str(len(hosts))}台主机，成功{success_count}，错误{error_count}台，命令执行错误{command_error}台\n')

    return True


# 清理屏幕
def clear_screen():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


# 键盘阻塞
def block_keyboard():
    if os.name == 'nt':
        while True:
            if msvcrt.kbhit():
                msvcrt.getch()
                # 忽略任何键盘输入
                # 可以根据需要进行逻辑处理
                continue
            break


# 输入错误
def invalid_input():
    clear_screen()
    block_keyboard()
    print('输入有误请重新输入')
    time.sleep(1)


# 获取键盘输入
def get_input():
    if os.name == 'nt':
        key = msvcrt.getch().decode('utf-8')
        return key
    else:
        file_descriptor = sys.stdin.fileno()
        old_settings = termios.tcgetattr(file_descriptor)
        try:
            tty.setraw(file_descriptor)
            key = sys.stdin.read(1)
        finally:
            termios.tcsetattr(file_descriptor, termios.TCSADRAIN, old_settings)
        return key


# 选择文件读取模式
def select_file_mode():
    excel_file_path = 'SSH.xlsx'
    sheet = 'Sheet1'
    result = False
    while True:
        clear_screen()
        print('请选择模式：')
        print('  1.txt模式')
        print('  2.excel模式')
        key = get_input()

        if key == '1':
            clear_screen()
            print('txt 模式')
            txt()
            break

        elif key == '2':
            clear_screen()
            while True:
                clear_screen()
                print(f'excel模式：')
                print(f'  1.文件名：{excel_file_path}')
                print(f'  2.表名：{sheet}')
                print(f'  3.继续')
                print(f'  0.返回上一层')
                key = get_input()

                if key == '1':
                    clear_screen()
                    new_file = input(f'当前文件名 {excel_file_path}，请输入新的文件名：\n')
                    if new_file:
                        excel_file_path = new_file
                elif key == '2':
                    clear_screen()
                    new_sheet = input(f'当前表名 {sheet}，请输入新的表名：\n')
                    if new_sheet:
                        sheet = new_sheet
                elif key == '3':
                    clear_screen()
                    try:
                        result = excel(excel_file_path, sheet)
                    except Exception as e:
                        print('')
                        print(e)
                        print('')
                        print('excel文件错误，按任意键返回。按esc退出...')
                        print('')
                        key = get_input()
                        if key == chr(27):  # chr(27) is esc
                            clear_screen()
                            sys.exit(1)
                        else:
                            continue
                    break
                elif key == '0':
                    break
                elif key == chr(27):  # chr(27)=esc
                    break
                else:
                    invalid_input()

            if result:
                break

        elif key == chr(27):  # chr(27)=esc
            clear_screen()
            sys.exit()

        else:
            invalid_input()

    block_keyboard()
    print("完成！按任意键退出...")
    get_input()


if __name__ == "__main__":
    select_file_mode()
