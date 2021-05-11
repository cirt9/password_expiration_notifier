from subprocess import Popen, PIPE
import re
from datetime import datetime, timedelta
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
import xmltodict
import logging
import logging.handlers


def get_list_of_users_to_notify(group_name):
    command = f'net group /domain "{group_name}"'
    pipe = Popen(command,shell=True,stdout=PIPE,stderr=PIPE)
    usernames_data = []
    save_line = False

    while True:
        line = pipe.stdout.readline()
        line_str = str(line)
        
        if not line or 'Polecenie zosta' in line_str:
            break

        if save_line:
            line_str = line_str.replace(r"b'", '')
            line_str = line_str.replace(r"'", '')
            line_str = line_str.replace(r'\r\n', '')

            usernames_data.append(line_str)
        else:
            if '---------------------------------------' in line_str:
                save_line = True
    
    return extract_usernames_from_usernames_data(usernames_data)


def extract_usernames_from_usernames_data(usernames_data):
    usernames = []

    for line in usernames_data:
        usernames_in_line = line.split(' ')
        usernames_in_line = [x for x in usernames_in_line if x != '']

        usernames += usernames_in_line
    
    return usernames


def get_password_expiration_date(username):    
    command = f'net user /domain {username}'
    pipe = Popen(command,shell=True,stdout=PIPE,stderr=PIPE)

    while True:
        line = pipe.stdout.readline()
        line_str = str(line)
        
        if 'wygasa' in line_str:
            return parse_date(line_str)
        if not line:
            break
    
    return None


def parse_date(text):
    match = re.search(r'\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}', text)
    return datetime.strptime(match.group(), '%Y-%m-%d %H:%M:%S') if match else None


def get_receiver_email(users, username):
    for user in users:
        if user['username'] == username:
            return user['email']
        
    return None


def run_notifier(config_filename, logger):
    with open(config_filename, encoding='utf-8') as config_file:
        config = xmltodict.parse(config_file.read(), force_list={'user', })

        notification_group_name = config['config']['notification_group_name']
        days_to_notify_in_advance = int(config['config']['days_to_notify_in_advance'])
        smtp_port = config['config']['smtp_port']
        smtp_server = config['config']['smtp_server']
        sender_email = config['config']['sender_email']
        sender_password = config['config']['sender_password']
        service_notification_email = config['config']['service_notification_email']
        password_will_expire_notification_subject = config['config']['password_will_expire_notification_subject']
        password_will_expire_notification_message = config['config']['password_will_expire_notification_message']
        password_expired_notification_subject = config['config']['password_expired_notification_subject']
        password_expired_notification_message = config['config']['password_expired_notification_message']
        service_notification_email_subject = config['config']['service_notification_email_subject']
        service_notification_email_message = config['config']['service_notification_email_message']
        user_email_not_found_message = config['config']['user_email_not_found_message']
        user_password_never_expires_message = config['config']['user_password_never_expires_message']
        users = config['config']['users']['user']

        usernames = get_list_of_users_to_notify(notification_group_name)
        context = ssl.create_default_context()

        with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
            server.login(sender_email, sender_password)

            service_notification_message = ''

            for username in usernames:
                password_expiration_date = get_password_expiration_date(username)

                if password_expiration_date:
                    today = datetime.now()
                    days_to_expire = (password_expiration_date - today).days
                    receiver_email = get_receiver_email(users, username)

                    if not receiver_email:
                        message = user_email_not_found_message.replace('{{ username }}', username) + '\n'
                        service_notification_message += message
                        logger.info(f'{username} email not found in config file.')
                    else:
                        if days_to_expire == -1:
                            subject = password_expired_notification_subject
                            message = password_expired_notification_message.replace('{{ username }}', username)
                            msg = MIMEText(message, _charset='UTF-8')
                            msg['Subject'] = Header(subject, 'utf-8')
                            msg['From'] = sender_email
                            msg['To'] = receiver_email

                            server.sendmail(sender_email, receiver_email, msg.as_string())
                            logger.info(f'{username} password has expired. Notification email was '\
                                f'sent to {receiver_email}')
                        elif 0 <= days_to_expire < days_to_notify_in_advance:
                            subject = password_will_expire_notification_subject.replace(
                                    '{{ date }}', password_expiration_date.strftime('%Y-%m-%d')).replace(
                                    '{{ time }}', password_expiration_date.strftime('%H:%M'))

                            message = password_will_expire_notification_message.replace(
                                    '{{ username }}', username).replace(
                                    '{{ date }}', password_expiration_date.strftime('%Y-%m-%d')).replace(
                                    '{{ time }}', password_expiration_date.strftime('%H:%M'))

                            msg = MIMEText(message, _charset='UTF-8')
                            msg['Subject'] = Header(subject, 'utf-8')
                            msg['From'] = sender_email
                            msg['To'] = receiver_email

                            server.sendmail(sender_email, receiver_email, msg.as_string())
                            logger.info(f'{username} password will expire on '\
                                    f'{password_expiration_date.strftime("%Y-%m-%d %H:%M")}. '\
                                    f'Notification email was sent to {receiver_email}')
                else:
                    message = user_password_never_expires_message.replace('{{ username }}', username) + '\n'
                    service_notification_message += message
                    logger.info(f'{username} password never expires.')

            if service_notification_message:
                subject = service_notification_email_subject
                message = service_notification_email_message.replace('{{ errors_list }}', service_notification_message)
                msg = MIMEText(message, _charset='UTF-8')
                msg['Subject'] = Header(subject, 'utf-8')
                msg['From'] = sender_email
                msg['To'] = service_notification_email

                server.sendmail(sender_email, service_notification_email, msg.as_string())
                logger.info(f'The errors in notifying process were found. Email to service '\
                            f'{service_notification_email} was sent.')


def main():
    log_filename = 'logs/logs.log'
    config_filename = 'config/config.xml'

    logger = logging.getLogger('logger')
    logger.setLevel(logging.INFO)
    handler = logging.handlers.RotatingFileHandler(log_filename, maxBytes=1000000, backupCount=1)
    formatter = logging.Formatter('%(asctime)s - %(message)s', '%Y-%m-%d %H:%M:%S')
    handler.setFormatter(formatter)
    logger.addHandler(handler)

    logger.info('----------Notifying has started----------')

    try:
        run_notifier(config_filename, logger)
    except Exception as e:
        logger.error(f'An error occured: {repr(e)}')
    else:
        logger.info('----------Notifying has ended successfully----------')


if __name__ == "__main__":
    main()
