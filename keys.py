from pynput.keyboard import Key, Listener, KeyCode

class Keys():
    def __init__(self):
        self.current_key = [0, 0]
        self.arrows = [Key.right, Key.left, Key.up, Key.down]
        self.new_key = None
        self.ended_key = []

    def set_ended_key(self):
        """ sets the list of buttons on the keyboard that you can finish work on """
        for key in list(Key):
            self.ended_key.append(key)
        for i in range(97, 122):
            self.ended_key.append(KeyCode.from_char(chr(i)))
        for key in self.arrows:
            self.ended_key.remove(key)

    def on_press(self, key):
        """ checks if the button is pressed """
        self.new_key = key
        if key == Key.right:
            self.current_key[0] += 1
        elif key == Key.left:
            self.current_key[0] -= 1
        if key == Key.up:
            self.current_key[1] -= 1
        elif key == Key.down:
            self.current_key[1] += 1

    def on_release(self, key):
        """ checks if the button is released """
        if key in self.arrows or key in self.ended_key:
            return False

    def listen(self):
        """ collect events until released """
        with Listener(on_press=self.on_press, on_release=self.on_release) as listener:
            listener.join()
            return self.current_key
