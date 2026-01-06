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
from pytribeam.GUI.runner_util import ExperimentController, ExperimentState
from pytribeam.GUI.runner_util.ui_components import ControlPanel, StatusPanel


class MainApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        # Create core
        tk.Tk.__init__(self, *args, **kwargs)
        self.title("TriBeam Runner")

        # Initialize resources and config
        self.resources = AppResources.from_module_file(__file__)
        self.app_config = AppConfig.from_env()
        self.app_config.ensure_directories()

        # Set icons
        self.iconbitmap(self.resources.icon_path)

        # Set the taskbar icon (Windows only)
        if os.name == "nt":
            import ctypes

            myappid = "pytribeam.tribeamlayeredacquisition"
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            self.iconbitmap(self.resources.icon_path)

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

        # Add experiment controller
        self.experiment_controller = ExperimentController()

        # Register callbacks for UI updates
        self.experiment_controller.register_callback(
            "state_changed", self._on_experiment_state_changed
        )
        self.experiment_controller.register_callback(
            "experiment_started", self._on_experiment_started
        )
        self.experiment_controller.register_callback(
            "experiment_completed", self._on_experiment_completed
        )
        self.experiment_controller.register_callback(
            "experiment_stopped", self._on_experiment_stopped
        )
        self.experiment_controller.register_callback(
            "validation_failed", self._on_validation_failed
        )
        self.experiment_controller.register_callback(
            "detector_warning", self._on_detector_warning
        )

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
        """Create control panel using ControlPanel component."""
        self.control_panel = ControlPanel(
            self,
            theme=self.theme,
            resources=self.resources,
        )
        self.control_panel.grid(row=0, column=0, rowspan=2, sticky="nsew")

        # Connect callbacks
        self.control_panel.on_new_config = self.new_config
        self.control_panel.on_load_config = self.load_config
        self.control_panel.on_edit_config = self.edit_config
        self.control_panel.on_validate_config = self.validate_config
        self.control_panel.on_start_experiment = self.start_experiment
        self.control_panel.on_stop_step = self.stop_step
        self.control_panel.on_stop_slice = self.stop_slice
        self.control_panel.on_stop_now = self.stop_hard

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
        """Create status panel using StatusPanel component."""
        self.status_panel = StatusPanel(
            self,
            theme=self.theme,
        )
        self.status_panel.grid(row=1, column=1, sticky="nsew")

    def change_theme(self):
        """Change the theme of the app."""
        if self.theme.theme_type == "light":
            self.theme = ctk.Theme("dark")
        else:
            self.theme = ctk.Theme("light")

        # Update the root
        self.configure(bg=self.theme.bg)

        # Update the control and status panels
        self.control_panel.destroy()
        self.status_panel.destroy()
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
        starting_slice = self.control_panel.starting_slice_var.get()
        starting_step = self.control_panel.starting_step_var.get()
        self._update_experiment_info()
        self.control_panel.starting_slice_var.set(starting_slice)
        self.control_panel.starting_step_var.set(starting_step)

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
        self.control_panel.set_validation_status(
            False, "Configuration file is unvalidated"
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
            self.control_panel.set_validation_status(
                True, "Configuration file is valid"
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
            self.control_panel.set_validation_status(
                True, "Configuration file is valid"
            )
            if return_settings:
                return experiment_settings
            else:
                return
        except Exception as e:
            messagebox.showerror(
                "Invalid config file", f"The provided config file is invalid:\n{e}"
            )
            self.control_panel.set_validation_status(
                False, "Configuration file is invalid"
            )
            return

    # -------- Update GUI functions -------- #

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
        # Prepare config info dictionary for control panel
        config_info = {
            "total_slices": max_slice_num,
            "total_steps": num_steps,
            "slice_thickness": f"{slice_thickness} µm",
            "config_path": str(self.config_path),
            "exp_dir": str(exp_dir),
            "step_names": step_names,
        }
        self.control_panel.update_experiment_info(config_info)

    def _update_slice_info(self, slice_number):
        """Update the slice information in the GUI."""
        self.status_panel.current_slice_var.set(slice_number)
        self.control_panel.starting_slice_var.set(slice_number)

    def _update_step_info(self, step_name):
        """Update the step information in the GUI."""
        self.status_panel.current_step_var.set(step_name)
        self.control_panel.starting_step_var.set(step_name)

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
        self.control_panel.start_btn.config(**start_kwards[start])
        self.control_panel.stop_step_btn.config(**step_kwargs[step])
        self.control_panel.stop_slice_btn.config(**slice_kwargs[slice])
        self.control_panel.stop_now_btn.config(**hard_kwargs[hard])
        # Note: Config buttons are not exposed by ControlPanel,
        # so we'll skip updating them for now
        self.update_idletasks()

    def _reset_starting_positions(self):
        """Reset the starting slice and step to defaults."""
        self.control_panel.starting_slice_var.set(1)
        step_names = self.control_panel.starting_step_menu.options
        if step_names:
            self.control_panel.starting_step_var.set(step_names[0])

    # ------- Experiment control functions -------- #

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
        # Update controller with current config path
        if self.config_path:
            self.experiment_controller.set_config_path(self.config_path)

        # Get starting position
        starting_slice = self.control_panel.starting_slice_var.get()
        starting_step = self.control_panel.starting_step_var.get()

        # Start the experiment via the controller
        success = self.experiment_controller.start_experiment(
            starting_slice=starting_slice, starting_step=starting_step
        )

        if success:
            # Update the validation status
            self.control_panel.set_validation_status(True)
            self._update_exp_control_buttons(start="disabled", buttons="disabled")

    # -------- Callbacks for experiment controller -------- #

    def _on_experiment_state_changed(self, state: ExperimentState):
        """Handle state updates from controller."""
        # Update status panel with new state
        self.status_panel.update_state(state)

        # Force GUI update
        try:
            self.update_idletasks()
        except tk.TclError:
            pass

    def _on_experiment_started(self, settings, start_slice, start_step):
        """Handle experiment start."""
        pass
        # print(f"Starting experiment at slice {start_slice}, step {start_step}")

    def _on_experiment_completed(self):
        """Handle experiment completion."""
        print("-----> Experiment complete <-----")
        self._reset_starting_positions()
        self._update_exp_control_buttons()

    def _on_experiment_stopped(self, final_slice, final_step):
        """Handle experiment stop."""
        print("-----> Experiment stopped <-----")
        # Update starting positions for resume
        self.control_panel.starting_slice_var.set(final_slice)
        self.control_panel.starting_step_var.set(final_step)
        self._update_exp_control_buttons()

    def _on_validation_failed(self, error_message):
        """Handle validation failure."""
        messagebox.showerror("Invalid config", error_message)
        self._update_exp_control_buttons()

    def _on_detector_warning(self, warning_message):
        """Handle EBSD/EDS detector warning."""
        messagebox.showwarning("Warning", warning_message)

    def stop_step(self):
        """Request stop after step."""
        self.experiment_controller.request_stop_after_step()
        self._update_exp_control_buttons(
            start="disabled", step="disabled", slice="disabled", hard="normal"
        )

    def stop_slice(self):
        """Request stop after slice."""
        self.experiment_controller.request_stop_after_slice()
        self._update_exp_control_buttons(
            start="disabled", step="normal", slice="disabled", hard="normal"
        )

    def stop_hard(self):
        """Request immediate stop."""
        self.experiment_controller.request_stop_now()
        self._update_exp_control_buttons(
            start="disabled", step="disabled", slice="disabled", hard="disabled"
        )

    # -------- Legacy experiment control functions (now handled by ExperimentController) -------- #

    def start_experiment_old(self):
        # Set the start exp button to be disabled and green
        self._update_exp_control_buttons(start="disabled", buttons="disabled")

        # Grab experiment info
        starting_slice = self.control_panel.starting_slice_var.get()
        starting_step_name = self.control_panel.starting_step_var.get()

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
            if experiment_settings.enable_EBSD:
                message_part1 = "EDS is not enabled"
            elif experiment_settings.enable_EDS:
                message_part1 = "EBSD is not enabled"
            else:
                message_part1 = "EBSD and EDS are not enabled"

            message_part2 = ", you will not have access to safety checking and these modalities during data collection. Please ensure these detectors are retracted before proceeding."
            messagebox.showwarning("Warning", message_part1 + message_part2)

        # Setup the progress bar
        start_point = (starting_slice - 1) * num_steps + starting_step_number
        self.status_panel.progress.set(
            int((start_point - 1) / (ending_slice * num_steps) * 100)
        )
        self.status_panel.current_step_var.set(step_names[starting_step_number])
        self.status_panel.current_slice_var.set(starting_slice)

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
                        self.status_panel.progress.set(perc_done)
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
                    self.status_panel.slice_time_var.set(avg_time_str)
                    self.status_panel.time_left_var.set(remaining_time_str)

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
                self.control_panel.starting_slice_var.set(i + 1)
                self.control_panel.starting_step_var.set(step_names[0])
            elif not stop_now:
                self.control_panel.starting_slice_var.set(i)
                self.control_panel.starting_step_var.set(step_names[j + 1])
            else:
                pass
        elif i == ending_slice and j == num_steps - 1:
            # The experiment finished
            print("-----> Experiment complete <-----")
            self.control_panel.starting_slice_var.set(1)
            self.control_panel.starting_step_var.set(step_names[0])
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

    def stop_step_old(self):
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

    def stop_slice_old(self):
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

    def stop_hard_old(self):
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
