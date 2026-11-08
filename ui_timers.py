"""Tk callbacks owned by one window, cancelled when that window is destroyed."""
import tkinter as tk


class WindowTimers:
    def __init__(self,window):
        self.window=window
        self.pending=set()
        self.closed=False
        window.bind('<Destroy>',self.destroyed,add='+')

    def after(self,delay,callback):
        if self.closed:return None
        def call():
            self.pending.discard(job)
            if not self.closed:callback()
        job=self.window.after(delay,call)
        self.pending.add(job)
        return job

    def destroyed(self,event):
        if event.widget is not self.window:return
        self.closed=True
        for job in tuple(self.pending):
            try:self.window.after_cancel(job)
            except tk.TclError:pass
        self.pending.clear()
