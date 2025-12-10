import sys
import shutil
import time
import datetime
import os
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
from pathlib import Path
from PIL import Image, ImageTk
import contextlib
import traceback

# pytribeam imports
import pytribeam.GUI.CustomTkinterWidgets as ctk
from pytribeam.GUI.config_ui.App import Configurator
from pytribeam import workflow, stage, utilities, log, laser, insertable_devices
import pytribeam.types as tbt

# Import refactored common utilities
from pytribeam.GUI.common import (
    AppResources,
    AppConfig,
    StoppableThread,
    TextRedirector,
)
from pytribeam.GUI.common.threading_utils import generate_escape_keypress


class MainApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        # Create core
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("TriBeam Runner")

        # Initialize resources and config
        self.resources = AppResources.from_module_file(__file__)
        self.app_config = AppConfig.from_env()
        self.app_config.ensure_directories()

        # Get images
        self.iconbitmap(self.resources.icon_path)
        self.image = Image.open(self.resources.logo_dark_path)
        self.image_size = (self.image.size[0] // 3, self.image.size[1] // 3)
        self.image.thumbnail(self.image_size, Image.ANTIALIAS)
        self.image = ImageTk.PhotoImage(self.image)

        # Set the window size
        self.frame_w = int(1200)
        self.frame_h = int(620)
        self.geometry(f"{self.frame_w}x{self.frame_h}")
        self.resizable(False, False)

        # Set the grid structure
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=4)
        self.grid_rowconfigure(1, weight=1)

        # Create variables
        self.config_path = None
        self.stop_after_slice = tk.BooleanVar(False)
        self.stop_after_step = tk.BooleanVar(False)
        self.stop_now = tk.BooleanVar(False)
        self.starting_slice_var = tk.IntVar()
        self.starting_step_var = tk.StringVar()
        self.starting_slice_var.set(1)
        self.starting_step_var.set("-")
        self.yml_version = None

        # Set the theme
        self.theme = ctk.Theme("dark")
        self.configure(bg=self.theme.bg)
        self._draw()

        # Map stdout through decorator so that we get the stdout in the GUI and CLI
        self.original_out = sys.stdout
        self.original_err = sys.stderr
        self.terminal_log_path = self.app_config.get_terminal_log_path()
        sys.stdout = TextRedirector(
            self.terminal, tag="stdout", log_path=str(self.terminal_log_path)
        )
        sys.stderr = TextRedirector(
            self.terminal, tag="stderr", log_path=str(self.terminal_log_path)
        )
        print("---")
        print("Welcome to TriBeam Layered Acquisition!")
        print("Please create a new configuration file or load an existing one.")
        print("---")

        # Bind the close button to the close function
        self.protocol("WM_DELETE_WINDOW", self.quit)
        self.thread_obj = None

        # Bind Ctrl+Shift+X to stop after step and Ctrl+X to stop after slice
        self.bind("<Control-X>", lambda e: self.stop_slice())
        self.bind("<Control-Shift-X>", lambda e: self.stop_step())

    def _draw(self):
        self._create_menubar()
        self._create_control_frame()
        self._create_display_frame()
        self._create_status_frame()

    def _create_menubar(self):
        # Create the menubar
        self.menu = tk.Menu(self)
        self.menu.add_command(label="Help", command=self.open_help)
        self.menu.add_command(label="Clear terminal", command=self.clear_terminal)
        self.menu.add_command(label="Test connections", command=self.test_connections)
        self.menu.add_command(label="Export log", command=self.export_log)
        self.menu.add_command(label="Change theme", command=self.change_theme)
        self.menu.add_command(label="Exit", command=self.quit)
        self.config(menu=self.menu)

    def _create_control_frame(self):
        self.control_frame = tk.Frame(self, bg=self.theme.bg, relief="ridge", bd=2)
        self.control_frame.grid(row=0, column=0, rowspan=2, sticky="nsew")
        self.control_frame.columnconfigure([0, 1, 2, 3], weight=1)
        l = tk.Label(
            self.control_frame,
            font=ctk.HEADER_FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
            image=self.image,
            width=self.image_size[0],
            height=self.image_size[1],
        )
        l.grid(row=0, column=0, columnspan=4)

        # Create experiment info labelframe
        sub_frame = tk.LabelFrame(
            self.control_frame,
            text="Experiment info",
            font=ctk.SUBHEADER_FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
        )
        sub_frame.grid(row=1, column=0, columnspan=4, sticky="nsew", pady=5, padx=5)
        sub_frame.columnconfigure(0, weight=5)
        sub_frame.columnconfigure(1, weight=1)
        sub_frame.rowconfigure([0, 1, 2, 3, 4, 5], weight=1)
        self.total_slices_l = tk.Label(
            sub_frame,
            text="Total number of slices: -",
            font=ctk.FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="w",
        )
        self.total_slices_l.grid(row=0, column=0, sticky="nsew", pady=2, padx=2)
        self.total_steps_l = tk.Label(
            sub_frame,
            text="Number of steps per slice: -",
            font=ctk.FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="w",
        )
        self.total_steps_l.grid(row=1, column=0, sticky="nsew", pady=2, padx=2)
        self.slice_thickness_l = tk.Label(
            sub_frame,
            text="Slice thickness: -",
            font=ctk.FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="w",
        )
        self.slice_thickness_l.grid(row=2, column=0, sticky="nsew", pady=2, padx=2)
        self.config_file_path = tk.Label(
            sub_frame,
            text="No configuration file loaded",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="w",
        )
        self.config_file_path.grid(
            row=3, column=0, columnspan=2, sticky="nsew", pady=2, padx=2
        )
        self.exp_dir_l = tk.Label(
            sub_frame,
            text="Exp dir: -",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="w",
        )
        self.exp_dir_l.grid(
            row=4, column=0, columnspan=2, sticky="nsew", pady=2, padx=2
        )
        self.valid_status = tk.Label(
            sub_frame,
            text="...",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="w",
        )
        self.valid_status.grid(
            row=5, column=0, columnspan=2, sticky="nsew", pady=2, padx=2
        )
        self.create_new_config_button = tk.Button(
            sub_frame,
            text="Create",
            font=ctk.FONT,
            command=self.new_config,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.create_new_config_button.grid(
            row=0, column=1, sticky="nsew", pady=2, padx=2
        )
        self.load_config_button = tk.Button(
            sub_frame,
            text="Load",
            font=ctk.FONT,
            command=self.load_config,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.load_config_button.grid(row=1, column=1, sticky="nsew", pady=2, padx=2)
        self.edit_config_button = tk.Button(
            sub_frame,
            text="Edit",
            font=ctk.FONT,
            command=self.edit_config,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.edit_config_button.grid(row=2, column=1, sticky="nsew", pady=2, padx=2)
        self.validate_config_button = tk.Button(
            sub_frame,
            text="Validate",
            font=ctk.FONT,
            command=self.validate_config,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.validate_config_button.grid(
            row=5, column=1, columnspan=2, sticky="nsew", pady=2, padx=2
        )

        # Put in the configuration buttons
        l = tk.Label(
            self.control_frame,
            text="Starting slice",
            font=ctk.FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="e",
        )
        l.grid(row=2, column=0, sticky="nsew", pady=5, padx=5)
        self.starting_slice = tk.Spinbox(
            self.control_frame,
            font=ctk.FONT,
            width=4,
            from_=0,
            to=10,
            bg=self.theme.bg_off,
            buttonbackground=self.theme.bg_off,
            disabledbackground=self.theme.bg_off,
            fg=self.theme.fg,
            textvariable=self.starting_slice_var,
            state="disabled",
        )
        self.starting_slice.grid(row=2, column=1, sticky="nsew", pady=5, padx=5)
        l = tk.Label(
            self.control_frame,
            text="Starting step",
            font=ctk.FONT,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="e",
        )
        l.grid(row=2, column=2, sticky="nsew", pady=5, padx=5)
        self.starting_step = ctk.MenuButton(
            self.control_frame,
            font=ctk.FONT,
            options=["-"],
            var=self.starting_step_var,
            width=12,
            state="disabled",
            bg=self.theme.bg_off,
            fg=self.theme.fg,
            h_bg=self.theme.accent1,
            h_fg=self.theme.accent1_fg,
        )
        self.starting_step.grid(row=2, column=3, sticky="nsew", pady=5, padx=5)

        # Put in the control buttons
        self.start_exp_b = tk.Button(
            self.control_frame,
            text="Start experiment",
            font=ctk.FONT,
            command=self.start_experiment,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.start_exp_b.grid(
            row=3, column=0, columnspan=4, sticky="nsew", pady=3, padx=5
        )
        self.stop_step_b = tk.Button(
            self.control_frame,
            text="Stop after current step",
            font=ctk.FONT,
            command=self.stop_step,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.stop_step_b.grid(
            row=4, column=0, columnspan=4, sticky="nsew", pady=3, padx=5
        )
        self.stop_slice_b = tk.Button(
            self.control_frame,
            text="Stop after current slice",
            font=ctk.FONT,
            command=self.stop_slice,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        self.stop_slice_b.grid(
            row=5, column=0, columnspan=4, sticky="nsew", pady=3, padx=5
        )
        self.stop_now_b = tk.Button(
            self.control_frame,
            text="Hard stop",
            font=ctk.FONT,
            command=self.stop_hard,
            bg=self.theme.bg_off,
            fg=self.theme.fg,
        )
        sep = tk.Frame(self.control_frame, bg=self.theme.bg, height=5, relief="flat")
        sep.grid(row=6, column=0, columnspan=4, sticky="nsew", pady=10)
        self.stop_now_b.grid(
            row=7, column=0, columnspan=4, sticky="nsew", pady=3, padx=5
        )

        # Put on tooltips
        ctk.tooltip(self.create_new_config_button, "Create a new configuration file")
        ctk.tooltip(self.load_config_button, "Load an existing configuration file")
        ctk.tooltip(self.edit_config_button, "Edit the current configuration file")
        ctk.tooltip(self.starting_slice, "Slice number to start the experiment at.")
        ctk.tooltip(self.starting_step, "Step number to start the experiment at.")
        ctk.tooltip(
            self.start_exp_b,
            "Start the experiment with the current configuration at the selected slice and step.",
        )
        ctk.tooltip(
            self.stop_step_b,
            "Stop the experiment after the current step is complete. (Ctrl+Shift+X)",
        )
        ctk.tooltip(
            self.stop_slice_b,
            "Stop the experiment after the current slice is complete. (Ctrl+X)",
        )
        ctk.tooltip(self.stop_now_b, "Stop the experiment immediately. (Ctrl+C)")

    def _create_display_frame(self):
        self.display_frame = tk.Frame(self, bg=self.theme.bg)
        self.display_frame.grid(row=0, column=1, sticky="nsew")
        self.display_frame.columnconfigure(0, weight=1)
        self.display_frame.rowconfigure(0, weight=1)
        # Put a listbox in the display frame that will act as a terminal for the program
        self.terminal = ctk.ScrolledText(
            self.display_frame,
            hscroll=False,
            bg=self.theme.terminal,
            fg=self.theme.terminal_fg,
            sbar_bg=self.theme.terminal,
            sbar_fg=self.theme.bg,
            autoscroll=True,
            font=ctk.FONT,
            wrap="word",
            state="disabled",
            bd=1,
        )
        self.terminal.grid(row=0, column=0, sticky="nsew")
        # Update the terminal by just changing the colors of the scrollbar
        style = self.terminal.vbar.make_style(
            orient="vertical",
            troughcolor=self.theme.terminal,
            background=self.theme.scrollbar,
            arrowcolor=self.theme.terminal,
        )
        self.terminal.vbar.config(style=style)

    def _create_status_frame(self):
        self.status_frame = tk.Frame(self, bg=self.theme.bg, relief="ridge", bd=2)
        self.status_frame.grid(row=1, column=1, sticky="nsew")
        self.status_frame.columnconfigure([0, 1, 2, 3, 4, 5, 6, 7], weight=1)
        self.status_frame.rowconfigure([0, 1], weight=1)
        # Current step
        current_step_label = tk.Label(
            self.status_frame,
            text="Current step",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="e",
        )
        current_step_label.grid(row=0, column=0, sticky="nsew", pady=5, padx=5)
        self.current_step = tk.StringVar()
        self.current_step.set("-")
        current_step_label2 = tk.Label(
            self.status_frame,
            textvariable=self.current_step,
            font=ctk.FONT_BOLD,
            bg=self.theme.bg,
            fg=self.theme.accent2,
            anchor="w",
        )
        current_step_label2.grid(row=0, column=1, sticky="nsew", pady=5, padx=5)
        # Current slice
        current_slice_label = tk.Label(
            self.status_frame,
            text="Current slice",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="e",
        )
        current_slice_label.grid(row=0, column=2, sticky="nsew", pady=5, padx=5)
        self.current_slice = tk.StringVar()
        self.current_slice.set("-")
        current_slice_label2 = tk.Label(
            self.status_frame,
            textvariable=self.current_slice,
            font=ctk.FONT_BOLD,
            bg=self.theme.bg,
            fg=self.theme.accent2,
            anchor="w",
        )
        current_slice_label2.grid(row=0, column=3, sticky="nsew", pady=5, padx=5)
        # Average slice time
        slice_time_label = tk.Label(
            self.status_frame,
            text="Average slice time",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="e",
        )
        slice_time_label.grid(row=0, column=4, sticky="nsew", pady=5, padx=5)
        self.slice_time = tk.StringVar()
        self.slice_time.set("-")
        slice_time_label2 = tk.Label(
            self.status_frame,
            textvariable=self.slice_time,
            font=ctk.FONT_BOLD,
            bg=self.theme.bg,
            fg=self.theme.accent2,
            anchor="w",
        )
        slice_time_label2.grid(row=0, column=5, sticky="nsew", pady=5, padx=5)
        # Time left label
        time_left_label = tk.Label(
            self.status_frame,
            text="Remaining duration",
            font=ctk.FONT_ITALIC,
            bg=self.theme.bg,
            fg=self.theme.fg,
            anchor="e",
        )
        time_left_label.grid(row=0, column=6, sticky="nsew", pady=5, padx=5)
        self.time_left = tk.StringVar()
        self.time_left.set("-")
        time_left_label2 = tk.Label(
            self.status_frame,
            textvariable=self.time_left,
            font=ctk.FONT_BOLD,
            bg=self.theme.bg,
            fg=self.theme.accent2,
            anchor="w",
        )
        time_left_label2.grid(row=0, column=7, sticky="nsew", pady=5, padx=5)
        # Progress bar
        self.progress = ctk.Progressbar(
            self.status_frame,
            bg=self.theme.bg,
            fg=self.theme.green,
            text_fg=self.theme.fg,
            text_bg=self.theme.bg,
        )
        self.progress.grid(row=1, column=0, columnspan=8, sticky="nsew", pady=5, padx=5)

    def change_theme(self):
        """Change the theme of the app."""
        if self.theme.theme_type == "light":
            self.theme = ctk.Theme("dark")
            img_path = self.resources.logo_dark_path
        else:
            self.theme = ctk.Theme("light")
            img_path = self.resources.logo_light_path
        # Update the root
        self.configure(bg=self.theme.bg)

        # Update the image
        self.image = Image.open(img_path)
        self.image_size = (self.image.size[0] // 3, self.image.size[1] // 3)
        self.image.thumbnail(self.image_size, Image.ANTIALIAS)
        self.image = ImageTk.PhotoImage(self.image)

        # Update the control and status frames
        self.control_frame.destroy()
        self.status_frame.destroy()
        self._create_control_frame()
        self._create_status_frame()

        # Update the terminal
        self.terminal.config(bg=self.theme.terminal, fg=self.theme.terminal_fg)
        style = self.terminal.vbar.make_style(
            orient="vertical",
            troughcolor=self.theme.terminal,
            background=self.theme.scrollbar,
            arrowcolor=self.theme.terminal,
        )
        self.terminal.vbar.config(style=style)

        # Make sure the experiment info is still in tact
        starting_slice = self.starting_slice_var.get()
        starting_step = self.starting_step_var.get()
        self._update_experiment_info()
        self.starting_slice_var.set(starting_slice)
        self.starting_step_var.set(starting_step)

    def test_connections(self):
        """Test the connections to the EDS/EBSD and the laser."""
        with WaitCursor(self):
            out_dict = {"result": None}
            self.thread_obj = StoppableThread(
                target=wrapper_for_output, args=(laser._device_connections, out_dict)
            )
            self.thread_obj.start()
            while self.thread_obj.is_alive():
                try:
                    self.update()
                except tk.TclError:
                    return
        status = out_dict["result"]
        messagebox.showinfo("Connection status", str(status))

    def clear_terminal(self):
        self.terminal.config(state=tk.NORMAL)
        self.terminal.delete("1.0", tk.END)
        self.terminal.config(state=tk.DISABLED)
        print("---")
        print("Welcome to TriBeam Layered Acquisition!")
        print("Please create a new configuration file or load an existing one.")
        print("---")

    def export_log(self):
        """Export the log file to a user-selected location."""
        if not self.terminal_log_path:
            messagebox.showerror("Error", "No log file to export.")
            return
        save_path = Path(
            filedialog.asksaveasfilename(
                title="Save log file",
                filetypes=[("Text files", "*.txt")],
                initialdir=os.getcwd(),
            )
        )
        if save_path == Path():
            return
        if not save_path.suffix == ".txt":
            save_path = save_path.with_suffix(".txt")
        shutil.copy(self.terminal_log_path, save_path)
        messagebox.showinfo("Success", f"Log file saved to {save_path}")

    def new_config(self):
        print("Creating new configuration file")
        self.edit_config(new=True)

    def load_config(self):
        self.config_path = Path(
            filedialog.askopenfilename(
                title="Select a configuration file",
                filetypes=[("YAML files", ("*.yaml", "*.yml"))],
                initialdir=os.getcwd(),
            )
        )
        if not self.config_path.is_file():
            print("No file selected.")
            return
        print(f"Imported configuration file from: {self.config_path}")
        self.valid_status.config(
            text="Configuration file is unvalidated", fg=self.theme.red
        )
        self._update_experiment_info()

    def edit_config(self, new=False):
        if new:
            app = Configurator(self, theme=self.theme)
        else:
            app = Configurator(self, theme=self.theme, yml_path=self.config_path)
        self.wait_window(app.toplevel)
        if app.clean_exit:
            self.config_path = Path(app.YAML_PATH)
            print(f"Imported configuration file from: {self.config_path}")
            self.valid_status.config(
                text="Configuration file is valid", fg=self.theme.green
            )
            self._update_experiment_info()

    def validate_config(self, return_settings=False):
        if self.config_path is None:
            messagebox.showerror("Error", "No configuration file loaded.")
            return
        try:
            with WaitCursor(self):
                out_dict = {"result": None, "error": False}
                self.thread_obj = StoppableThread(
                    target=wrapper_for_output,
                    args=(workflow.pre_flight_check, out_dict, self.config_path),
                )
                self.thread_obj.start()
                while self.thread_obj.is_alive():
                    try:
                        self.update()
                    except tk.TclError:
                        return
                    except Exception as e:
                        break
                if out_dict["error"]:
                    raise Exception(out_dict["error"])
            experiment_settings = out_dict["result"]
            self.valid_status.config(
                text="Configuration file is valid", fg=self.theme.green
            )
            if return_settings:
                return experiment_settings
            else:
                return
        except Exception as e:
            messagebox.showerror(
                "Invalid config file", f"The provided config file is invalid:\n{e}"
            )
            self.valid_status.config(
                text="Configuration file is invalid", fg=self.theme.red
            )
            return

    def _update_experiment_info(self):
        """Update the experiment information in the GUI from the current yaml file."""
        if self.config_path is None:
            return
        try:
            self.yml_version = utilities.yml_version(self.config_path)
            db = utilities.yml_to_dict(
                yml_path_file=self.config_path,
                version=self.yml_version,
                required_keys=("general", "steps"),
            )
            num_steps = db["general"]["step_count"]
            max_slice_num = db["general"]["max_slice_num"]
            slice_thickness = db["general"]["slice_thickness_um"]
            exp_dir = db["general"]["exp_dir"]
            step_names = list(db["steps"].keys())
        except Exception as e:
            messagebox.showerror("Error", f"Error loading configuration file:\n{e}")
            return
        # Split the config file path along a directory separator and only show the last 40 characters
        str_config_path = str(self.config_path).split(os.sep)
        while len(str_config_path) > 1 and len("".join(str_config_path)) > 30:
            str_config_path.pop(0)
        # Same for the experiment directory but 30 characters
        str_exp_dir = str(Path(exp_dir)).split(os.sep)
        while len(str_exp_dir) > 1 and len("".join(str_exp_dir)) > 30:
            str_exp_dir.pop(0)
        self.config_file_path.config(
            text=f"Config: ...{os.sep}{os.sep.join(str_config_path)}"
        )
        self.exp_dir_l.config(text=f"Exp dir: ...{os.sep}{os.sep.join(str_exp_dir)}")
        self.total_slices_l.config(text=f"Total number of slices: {max_slice_num}")
        self.total_steps_l.config(text=f"Number of steps per slice: {num_steps}")
        self.slice_thickness_l.config(text=f"Slice thickness: {slice_thickness}")
        self.starting_slice_var.set(1)
        self.starting_slice.config(from_=1, to=max_slice_num, state="normal")
        self.starting_step_var.set(step_names[0])
        self.starting_step.set_options(step_names)
        self.starting_step.config(state="normal")

    def _update_slice_info(self, slice_number):
        """Update the slice information in the GUI."""
        self.current_slice.set(slice_number)
        self.starting_slice_var.set(slice_number)

    def _update_step_info(self, step_name):
        """Update the step information in the GUI."""
        self.current_step.set(step_name)
        self.starting_step_var.set(step_name)

    def run_in_thread(self, func, *args, **kwargs):
        out_dict = {"error": False}
        self.thread_obj = StoppableThread(
            target=func, args=(out_dict,) + args, kwargs=kwargs
        )
        self.thread_obj.start()
        while self.thread_obj.is_alive():
            try:
                self.update()
            except tk.TclError:
                return
            if self.stop_now.get() and self.thread_obj.is_alive():
                generate_escape_keypress()
                self.thread_obj.raise_exception(KeyboardInterrupt)
                break
        return out_dict

    def start_experiment(self):
        """
        Start the experiment.

        This function is the main function that starts the experiment. It will loop over the
        slices and steps, calling the step function for each step. The starting slice and
        step are taken from the entries in the GUI.

        Stopping the experiment is control by keystrokes and the control buttons in the GUI.
        The experiment can be stopped after the current step, after the current slice, or
        immediately (hard stop). A modified Thread, ThreadWithExc, is used to run the steps
        in a separate thread so that the main thread can update the GUI and check for stop.
        The modified thread can raise a KeyboardInterrupt exception in the step function to
        stop the experiment.
        """
        # Set the start exp button to be disabled and green
        self._update_exp_control_buttons(start="disabled", buttons="disabled")

        # Grab experiment info
        starting_slice = self.starting_slice_var.get()
        starting_step_name = self.starting_step_var.get()

        # Run preflight check
        experiment_settings: tbt.ExperimentSettings = self.validate_config(
            return_settings=True
        )
        if experiment_settings is None:
            self._update_exp_control_buttons(start="normal", buttons="normal")
            return
        else:
            # Process the experiment settings
            num_steps = experiment_settings.general_settings.step_count
            step_names = [i.name for i in experiment_settings.step_sequence]
            starting_step_number = step_names.index(starting_step_name)
            ending_slice = experiment_settings.general_settings.max_slice_number
            # Log the experiment settings
            log.experiment_settings(
                slice_number=starting_slice,
                step_number=starting_step_number,
                log_filepath=experiment_settings.general_settings.log_filepath,
                yml_path=self.config_path,
            )
            print("Preflight check successful")

        # Check if EBSD and EDS are enabled
        if not experiment_settings.enable_EBSD or not experiment_settings.enable_EDS:
            pick = -1
            if not experiment_settings.enable_EBSD:
                pick += 1
            if not experiment_settings.enable_EDS:
                pick += 2
            message_part1 = [
                "EBSD is not enabled",
                "EDS is not enabled",
                "EBSD and EDS are not enabled",
            ][pick]
            message_part2 = ", you will not have access to safety checking and these modalities during data collection. Please ensure these detectors are retracted before proceeding."
            messagebox.showwarning("Warning", message_part1 + message_part2)

        # Setup the progress bar
        start_point = (starting_slice - 1) * num_steps + starting_step_number
        self.progress.set(int((start_point - 1) / (ending_slice * num_steps) * 100))
        self.current_step.set(step_names[starting_step_number])
        self.current_slice.set(starting_slice)

        # Setup timer
        slice_times = []

        # Run the experiment (loop over the slices and steps)
        # Three levels here: try to catch KeyboardInterrupt, for i in slices, for j in steps
        print(
            f'Starting experiment at slice {starting_slice} and step "{starting_step_name}", number {starting_step_number+1} of {len(step_names)}'
        )
        try:
            if self.stop_now.get():
                raise KeyboardInterrupt
            for i in range(starting_slice, ending_slice + 1):
                # Get the slice start time and update the slice info
                t0 = time.time()
                self._update_slice_info(i)

                for j in range(num_steps):
                    # Skip steps if we are starting in the middle of a slice
                    if i == starting_slice and j < starting_step_number:
                        continue

                    # Update the current step
                    self._update_step_info(step_names[j])

                    # Create a thread object to run the step_call function
                    args = (i, j + 1, experiment_settings)
                    out_dict = self.run_in_thread(step_call_wrapper, *args)
                    stop_step = self.stop_after_step.get()
                    stop_slice = self.stop_after_slice.get()
                    stop_now = self.stop_now.get()
                    if out_dict["error"]:
                        stop_now = True

                    # Update progress bar if we didn't hard stop
                    if not stop_now:
                        perc_done = int(
                            ((i - 1) * num_steps + (j + 1))
                            / (ending_slice * num_steps)
                            * 100
                        )
                        self.progress.set(perc_done)
                        try:
                            self.update_idletasks()
                        except tk.TclError:
                            return

                    # Break out if we are stopping this step or right now
                    if stop_step or stop_now:
                        break

                # Break out if we are stopping at all
                if stop_step or stop_slice or stop_now:
                    break
                else:
                    try:
                        self.update_idletasks()
                    except tk.TclError:
                        return
                    t1 = time.time()
                    slice_times.append(t1 - t0)
                    avg_time = round(sum(slice_times) / len(slice_times))
                    remaining_time = avg_time * (ending_slice - i)
                    avg_time_str = str(datetime.timedelta(seconds=avg_time))
                    remaining_time_str = str(datetime.timedelta(seconds=remaining_time))
                    self.slice_time.set(avg_time_str)
                    self.time_left.set(remaining_time_str)

        except KeyboardInterrupt:
            print(
                "-----> Experiment was stopped immediately by user (keyboard interrupt)"
            )
            if self.thread_obj is not None and self.thread_obj.is_alive():
                self.thread_obj.raise_exception(KeyboardInterrupt)
                stop_now = True

        # Handle the end of the experiment
        if stop_now or stop_step or (stop_slice and i != ending_slice):
            # The experiment did not finish
            print("-----> Experiment stopped <-----")
            if not stop_now and j + 1 == num_steps:
                self.starting_slice_var.set(i + 1)
                self.starting_step_var.set(step_names[0])
            elif not stop_now:
                self.starting_slice_var.set(i)
                self.starting_step_var.set(step_names[j + 1])
            else:
                pass
        elif i == ending_slice and j == num_steps - 1:
            # The experiment finished
            print("-----> Experiment complete <-----")
            self.starting_slice_var.set(1)
            self.starting_step_var.set(step_names[0])
        else:
            # The experiment ended for an unknown reason
            print("-----> Experiment stopped (unknown) <-----")
        # Ensure that all devices are retracted
        if not stop_now:
            insertable_devices.retract_all_devices(
                microscope=experiment_settings.microscope,
                enable_EBSD=experiment_settings.enable_EBSD,
                enable_EDS=experiment_settings.enable_EDS,
            )
        # Reset the stop flags
        self.stop_after_slice.set(False)
        self.stop_after_step.set(False)
        self.stop_now.set(False)
        # Update the GUI
        self._update_exp_control_buttons()

    def _update_exp_control_buttons(
        self,
        start="normal",
        step="normal",
        slice="normal",
        hard="normal",
        buttons="normal",
    ):
        """Update the experiment control buttons."""
        start_kwards = {
            "normal": {"state": "normal", "bg": self.theme.bg},
            "disabled": {
                "state": "disabled",
                "bg": self.theme.green,
                "disabledforeground": self.theme.bg,
            },
        }
        step_kwargs = {
            "normal": {"state": "normal", "bg": self.theme.bg},
            "disabled": {
                "state": "disabled",
                "bg": self.theme.accent3,
                "disabledforeground": self.theme.bg,
            },
        }
        slice_kwargs = {
            "normal": {"state": "normal", "bg": self.theme.bg},
            "disabled": {
                "state": "disabled",
                "bg": self.theme.accent3,
                "disabledforeground": self.theme.bg,
            },
        }
        hard_kwargs = {
            "normal": {"state": "normal", "bg": self.theme.bg},
            "disabled": {
                "state": "disabled",
                "bg": self.theme.accent3,
                "disabledforeground": self.theme.bg,
            },
        }
        self.start_exp_b.config(**start_kwards[start])
        self.stop_step_b.config(**step_kwargs[step])
        self.stop_slice_b.config(**slice_kwargs[slice])
        self.stop_now_b.config(**hard_kwargs[hard])
        self.create_new_config_button.config({"state": buttons})
        self.edit_config_button.config({"state": buttons})
        self.load_config_button.config({"state": buttons})
        self.update_idletasks()

    def stop_step(self):
        """
        Stop the experiment after the current step is complete.
        This is an experiment control function that sets a flag to stop the experiment after the current step is complete.
        """
        print("-----> Stopping after current step")
        if self.thread_obj is None:
            return
        self.stop_after_step.set(True)
        self._update_exp_control_buttons(
            start="disabled",
            step="disabled",
            slice="disabled",
            hard="normal",
            buttons="disabled",
        )

    def stop_slice(self):
        """
        Stop the experiment after the current slice is complete.
        This is an experiment control function that sets a flag to stop the experiment after the current slice is complete.
        """
        print("-----> Stopping after current slice")
        if self.thread_obj is None:
            return
        self.stop_after_slice.set(True)
        self._update_exp_control_buttons(
            start="disabled",
            step="normal",
            slice="disabled",
            hard="normal",
            buttons="disabled",
        )

    def stop_hard(self):
        """
        Stop the experiment immediately.
        This is an experiment control function that sets a flag to stop the experiment immediately.
        """
        print("-----> Experiment was stopped immediately by user (button press)")
        if self.thread_obj is None:
            return
        self.stop_now.set(True)
        self._update_exp_control_buttons(
            start="disabled",
            step="disabled",
            slice="disabled",
            hard="disabled",
            buttons="disabled",
        )

    def open_help(self):
        """Open the user guide in a web browser."""
        import webbrowser

        ### TODO: Fix this to open local file properly on all OSes
        # webbrowser.open(f"file://{self.resources.user_guide_path}")
        webbrowser.open(f"{self.resources.user_guide_path}")

    def quit(self):
        """
        Quit the program.
        This function is called when the user closes the window or selects the exit option from the menu.
        """
        sys.stdout = self.original_out
        sys.stderr = self.original_err
        if self.thread_obj is not None and self.thread_obj.is_alive():
            self.thread_obj.raise_exception(KeyboardInterrupt)
        self.update()
        self.destroy()


@contextlib.contextmanager
def WaitCursor(root):
    root.config(cursor="wait")
    root.update()
    try:
        yield root
    finally:
        root.config(cursor="")


def step_call_wrapper(out_dict, slice_number, step_index, experiment_settings):
    """
    A wrapper function to call the step function in a thread while also being able to catch a KeyboardInterrupt.
    If the exception is raised, the thread is stopped and the experiment is halted.
    """
    try:
        workflow.perform_step(slice_number, step_index, experiment_settings)
    except KeyboardInterrupt:
        try:
            stage.stop(experiment_settings.microscope)
            print("-----> Stage stop unsuccessful")
        except SystemError:
            print("-----> Stage stop successful")
        out_dict["error"] = True
        return False
    except Exception as e:
        print(
            f"Unexpected error in step {step_index} of slice {slice_number}: {e.__class__} {e}"
        )
        app_config = AppConfig.from_env()
        app_config.ensure_directories()
        err_path = app_config.get_error_log_path()
        with open(err_path, "w") as f:
            f.write(f"Exception: {type(e).__name__} - {e}\n")
            traceback.print_exc(file=f)
        try:
            stage.stop(experiment_settings.microscope)
            print("-----> Stage stop unsuccessful")
        except SystemError:
            print("-----> Stage stop successful")
        out_dict["error"] = True
        return False
    out_dict["error"] = False
    return True


def wrapper_for_output(func, out_dict, *args, **kwargs):
    try:
        out_dict["result"] = func(*args, **kwargs)
    except Exception as e:
        out_dict["error"] = e
    return out_dict


# Note: ThreadWithExc, TextRedirector, and generate_escape_keypress are now imported from common.threading_utils


if __name__ == "__main__":
    app = MainApplication()
    app.mainloop()
