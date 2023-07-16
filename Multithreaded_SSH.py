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
import excel_read

if os.name == 'nt':
    import msvcrt
else:
    import tty
    import termios

success_count = 0
error_count = 0
command_error = 0


# Command timeout error handling
class CommandTimeoutError(Exception):
    pass


class CustomError(Exception):
    pass


# Command execution module
def execute_command(shell, hostname, ssh_log_file, commands=None, ends=None, command_timeout=60):
    if commands is None:
        commands = {'cmd1': 'dis ver', 'cmd2': 'dis th'}
    if ends is None:
        ends = {}

    global command_error

    output = ''

    # Execute custom commands
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

    # Configure logging to record detailed paramiko logs (can be commented out if not needed)
    logging.basicConfig(level=logging.DEBUG, filename=f'log/paramiko-{datetime.date.today()}.log', filemode='w',
                        format='%(asctime)s - %(levelname)s - %(message)s')

    global success_count, error_count

    log_cache = []
    connected = False
    break_tag = False

    # Create log directory and file
    os.makedirs('log', exist_ok=True)
    os.makedirs('SSH_log', exist_ok=True)
    log_file = f'log/{datetime.date.today()}.txt'
    ssh_log_file = f'SSH_log/{datetime.date.today()}-{address}.txt'

    # Create SSH client
    client = paramiko.SSHClient()

    # Set up key verification directory and file
    # os.makedirs('SSH_key', exist_ok=True)
    # known_hosts_file = './SSH_key/known_hosts'
    # with open(known_hosts_file, 'w') as f:
    #     pass

    # local key verification
    # client.set_missing_host_key_policy(paramiko.WarningPolicy())
    # client.load_host_keys(os.path.expanduser(known_hosts_file))

    # Automatically add host keys
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    for password in passwords:
        try:
            print(f'{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]} Attempting connection to {address}')
            log_cache.append(
                f'{datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]} Attempting connection to {address} ')
            client.connect(hostname=address, port=port, username=username, password=password, timeout=timeout)

            connected = True
            print(f'Connected to {address}')
            log_cache.append(f'Connected ')

            # Create interactive shell
            shell = client.invoke_shell()

            time.sleep(0.5)
            output = ''

            # Get hostname
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
                    print(f'Error retrieving device name. Skipping {address}')
                    log_cache.append(f'Error retrieving device name. Skipping {address}\n')
                    break

            # Set temporary output length for output handling (tested only on Huawei devices!)
            shell.send(b'screen-length 0 temporary\n')  # Set output length to 0, temporary mode
            time.sleep(1)
            output += shell.recv(1024).decode('utf-8')

            # Write SSH log before executing custom commands (can be commented out if not needed)
            with open(ssh_log_file, 'a', encoding='utf-8') as f_sshlog:
                f_sshlog.write(output)

            # Execute custom commands
            execute = execute_command(shell, hostname, ssh_log_file, commands, ends, command_timeout)
            if execute:
                print(f'Execution completed')
                log_cache.append(f'Execution completed ')
                success_count += 1

        except paramiko.AuthenticationException:
            print(f'Authentication failed: {address}')
            log_cache.append(f'Authentication failed: {address}\n')

        except paramiko.SSHException as e:
            print(f'SSH connection error: {address}, {str(e)}')
            log_cache.append(f'SSH connection error: {address}, {str(e)}\n')
            break_tag = True

        except paramiko.ssh_exception.NoValidConnectionsError as e:
            print(f'Unable to connect to host: {address}, {str(e)}')
            log_cache.append(f'Unable to connect to host: {address}, {str(e)}\n')
            break_tag = True

        except TimeoutError as e:
            print(f'Connection timed out: {address}, {str(e)}')
            log_cache.append(f'Connection timed out: {address}, {str(e)}\n')
            break_tag = True

        except CommandTimeoutError as e:
            log_cache.append(f'Command execution timed out, skipping host: {address}, {str(e)}\n')

        except Exception as e:
            print(f'Unknown error: {address}, {str(e)}')
            log_cache.append(f'Encountered unknown error: {address}, {str(e)}\n')

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
                print(f'Connection closed: {address}')
                log_cache.append(f'Connection closed: {address}\n')
                with open(log_file, 'a', encoding='utf-8') as f_log:
                    for log in log_cache:
                        f_log.write(log)
                break

    # Error count
    if not connected:
        error_count += 1


def txt():
    # Create log directory and file
    os.makedirs('log', exist_ok=True)
    log_file = f'log/{datetime.date.today()}.txt'

    print(f'Program start time: {datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(f'Program start time: {datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]}\n')

    username = input('Enter username: ')
    if username == '':
        username = 'admin'

    # password = input('Enter device password (separated by "space", up to 3):').split(' ')
    passwords = getpass.getpass(prompt='Enter device password (separated by "space", up to 3):').split(' ')
    if passwords == ['']:
        passwords = ['admin@123']
    passwords = passwords[:3]

    # Read host address list
    with open('host.txt', 'r') as f:
        hosts = f.read().splitlines()

    # Read commands
    command_num = 1
    commands = {}
    with open('commands.txt', 'r') as f:
        cmds = f.read().splitlines()
    for cmd in cmds:
        commands['cmd' + str(command_num)] = cmd
        command_num += 1

    # Create threads
    threads = []
    for host in hosts:
        thread = threading.Thread(target=ssh, kwargs={'address': host, 'username': username, 'passwords': passwords,
                                                      'commands': commands})
        threads.append(thread)
        # time.sleep(0.2)
        thread.start()

    # Wait for threads to complete
    for thread in threads:
        thread.join()

    print(f'Execution completed. Total {str(len(hosts))} hosts, {success_count} successful, {error_count} errors, '
          f'{command_error} command timeout errors')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(
            f'Execution completed. Total {str(len(hosts))} hosts, {success_count} successful, {error_count} errors, '
            f'{command_error} command timeout errors\n')


def excel(excel_file='SSH.xlsx', sheet='Sheet1'):
    try:
        hosts, return_info = excel_read.process_excel_file(excel_file, sheet, True)
        for info in return_info:
            print(info)
        print(f'{str(len(hosts))} hosts')
    except Exception as e:
        raise CustomError(e)

    # Create log directory and file
    os.makedirs('log', exist_ok=True)
    log_file = f'log/{datetime.date.today()}.txt'

    # password = input('Enter device password (separated by "space", up to 3):').split(' ')
    passwords = getpass.getpass(prompt='Enter device password (separated by "space", up to 3):').split(' ')
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
        thread = threading.Thread(target=ssh,
                                  kwargs={'hostname': host['hostname'], 'address': host['Address'],
                                          'username': host['Username'], 'passwords': passwords, 'commands': commands,
                                          'ends': ends})
        threads.append(thread)
        # time.sleep(0.2)
        thread.start()

    for thread in threads:
        thread.join()

    print(f'Execution completed. Total {str(len(hosts))} hosts, {success_count} successful, {error_count} errors, '
          f'{command_error} command timeout errors')
    with open(log_file, 'a', encoding='utf-8') as f_log:
        f_log.write(
            f'Execution completed. Total {str(len(hosts))} hosts, {success_count}, {error_count} errors, '
            f'{command_error} command timeout errors\n')

    return True


# Clear the screen
def clear_screen():
    if os.name == 'nt':
        os.system('cls')
    else:
        os.system('clear')


# Keyboard blocking
def block_keyboard():
    if os.name == 'nt':
        while True:
            if msvcrt.kbhit():
                msvcrt.getch()
                # Ignore any keyboard input
                # Perform logic as required
                continue
            break


# Invalid input
def invalid_input():
    clear_screen()
    block_keyboard()
    print('Invalid input. Please try again.')
    time.sleep(1)


# Get keyboard input
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


# Select file reading mode
def select_file_mode():
    excel_file_path = 'SSH.xlsx'
    sheet = 'Sheet1'
    result = False
    while True:
        clear_screen()
        print('Select mode:')
        print('  1.txt mode')
        print('  2.excel mode')
        key = get_input()

        if key == '1':
            clear_screen()
            print('txt mode:')
            txt()
            break

        elif key == '2':
            clear_screen()
            while True:
                clear_screen()
                print(f'Excel mode:')
                print(f'  1.File: {excel_file_path}')
                print(f'  2.Sheet: {sheet}')
                print(f'  3.Continue')
                print(f'  0.Go back')
                key = get_input()

                if key == '1':
                    clear_screen()
                    new_file = input(f'Current file: {excel_file_path}. Enter a new file name:\n')
                    if new_file:
                        excel_file_path = new_file
                elif key == '2':
                    clear_screen()
                    new_sheet = input(f'Current sheet: {sheet}. Enter a new sheet name:\n')
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
                        print('Excel file error. Press any key to return. Press esc to exit...')
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
                elif key == chr(27):  # chr(27) is esc
                    break
                else:
                    invalid_input()

            if result:
                break

        elif key == chr(27):  # chr(27) is esc
            clear_screen()
            sys.exit()

        else:
            invalid_input()

    block_keyboard()
    print("Completed! Press any key to exit...")
    get_input()


if __name__ == "__main__":
    select_file_mode()
