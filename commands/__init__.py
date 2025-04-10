from .smart_post_dialog import entry as smart_post_dialog

commands = [
    smart_post_dialog,
]

def start():
    for command in commands:
        command.start()

def stop():
    for command in commands:
        command.stop()