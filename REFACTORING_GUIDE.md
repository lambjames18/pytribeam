# GUI Refactoring Guide

This document provides guidance for completing the GUI refactoring initiated in Phase 1-4. The refactoring separates concerns, improves testability, and makes the codebase more maintainable.

## Overview

The refactoring is organized into phases:
- ✅ **Phase 1**: Extract common utilities (completed)
- ✅ **Phase 2**: Extract models and business logic (completed)
- 🔄 **Phase 3**: Extract controllers (infrastructure complete, integration pending)
- 🔄 **Phase 4**: Refactor UI components (components created, integration pending)

## Completed Work

### Phase 1 & 2: Foundation (✅ Complete)

Created modular, well-documented components:

**Common utilities (`GUI/common/`):**
- `threading_utils.py`: Thread management, TextRedirector
- `resources.py`: Resource path management
- `errors.py`: Custom exception hierarchy
- `logging_config.py`: Logging configuration
- `config_manager.py`: Application settings

**Data models (`GUI/config_ui/`):**
- `pipeline_model.py`: PipelineConfig, StepConfig data structures
- `validator.py`: Configuration validation logic
- `microscope_interface.py`: Hardware abstraction layer

### Phase 3: Controllers (🔄 Infrastructure Complete)

**Runner controller (`GUI/runner/`):**
- `experiment_controller.py`: ExperimentController and ExperimentState
  - Manages experiment execution without UI dependencies
  - Provides callbacks for UI updates
  - Handles threading, stopping, progress tracking

**Config UI controller (`GUI/config_ui/`):**
- `editor_controller.py`: EditorController
  - Manages pipeline configuration state
  - Handles step selection, editing, validation
  - Decoupled from UI presentation

### Phase 4: UI Components (🔄 Components Created)

**Runner UI (`GUI/runner/`):**
- `ui_components.py`: ControlPanel, StatusPanel
  - Reusable, self-contained UI widgets
  - Callback-based integration
  - Theme-aware

## Integration Guide

### Integrating ExperimentController into runner.py

The `ExperimentController` is designed to replace the experiment logic currently embedded in `MainApplication`. Here's how to integrate it:

#### 1. Update MainApplication.__init__

```python
class MainApplication(tk.Tk):
    def __init__(self, *args, **kwargs):
        # ... existing setup ...

        # Add experiment controller
        self.experiment_controller = ExperimentController()

        # Register callbacks for UI updates
        self.experiment_controller.register_callback(
            'state_changed',
            self._on_experiment_state_changed
        )
        self.experiment_controller.register_callback(
            'experiment_started',
            self._on_experiment_started
        )
        self.experiment_controller.register_callback(
            'experiment_completed',
            self._on_experiment_completed
        )
        self.experiment_controller.register_callback(
            'experiment_stopped',
            self._on_experiment_stopped
        )
        self.experiment_controller.register_callback(
            'validation_failed',
            self._on_validation_failed
        )
```

#### 2. Replace start_experiment method

Replace the current ~175 line `start_experiment` method with:

```python
def start_experiment(self):
    """Start the experiment using controller."""
    # Update controller with current config
    if self.config_path:
        self.experiment_controller.set_config_path(self.config_path)

    # Get starting positions
    starting_slice = self.starting_slice_var.get()
    starting_step = self.starting_step_var.get()

    # Start via controller
    success = self.experiment_controller.start_experiment(
        starting_slice=starting_slice,
        starting_step=starting_step
    )

    if success:
        self._update_exp_control_buttons(start="disabled", buttons="disabled")
```

#### 3. Implement callback handlers

```python
def _on_experiment_state_changed(self, state: ExperimentState):
    """Handle state updates from controller."""
    # Update current slice/step displays
    self.current_slice.set(str(state.current_slice))
    self.current_step.set(state.current_step)

    # Update progress
    self.progress.set(state.progress_percent)

    # Update timing
    self.slice_time.set(state.avg_slice_time_str)
    self.time_left.set(state.remaining_time_str)

    # Update UI
    try:
        self.update_idletasks()
    except tk.TclError:
        pass

def _on_experiment_started(self, settings, start_slice, start_step):
    """Handle experiment start."""
    print(f"Starting experiment at slice {start_slice}, step {start_step}")

def _on_experiment_completed(self):
    """Handle experiment completion."""
    print("-----> Experiment complete <-----")
    self._reset_starting_positions()
    self._update_exp_control_buttons()

def _on_experiment_stopped(self, final_slice, final_step):
    """Handle experiment stop."""
    print("-----> Experiment stopped <-----")
    # Update starting positions for resume
    self.starting_slice_var.set(final_slice)
    # ... update step ...
    self._update_exp_control_buttons()

def _on_validation_failed(self, error_message):
    """Handle validation failure."""
    messagebox.showerror("Invalid config", error_message)
    self._update_exp_control_buttons()
```

#### 4. Simplify stop methods

```python
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
```

### Integrating UI Components

The `ControlPanel` and `StatusPanel` can gradually replace sections of `MainApplication`:

#### Option 1: Gradual Migration

Keep existing UI but use components internally:

```python
def _create_control_frame(self):
    """Create control frame using ControlPanel component."""
    self.control_panel = ControlPanel(
        self,
        theme=self.theme,
        resources=self.resources
    )
    self.control_panel.grid(row=0, column=0, rowspan=2, sticky="nsew")

    # Wire up callbacks
    self.control_panel.on_new_config = self.new_config
    self.control_panel.on_load_config = self.load_config
    self.control_panel.on_edit_config = self.edit_config
    # ... etc ...
```

#### Option 2: Complete Replacement

For a clean break, create new `MainWindow` class:

```python
# In new file: GUI/runner/main_window.py

class MainWindow(tk.Tk):
    """Main application window using refactored components."""

    def __init__(self):
        super().__init__()
        self.title("TriBeam Runner")

        # Initialize resources and config
        self.resources = AppResources.from_module_file(__file__)
        self.app_config = AppConfig.from_env()
        self.theme = ctk.Theme("dark")

        # Create controller
        self.controller = ExperimentController()
        self._register_controller_callbacks()

        # Create UI
        self._setup_window()
        self._create_ui()

        # Setup terminal redirection
        self._setup_terminal()

    def _create_ui(self):
        """Create UI using components."""
        # Control panel
        self.control_panel = ControlPanel(
            self, self.theme, self.resources
        )
        self.control_panel.grid(row=0, column=0, rowspan=2, sticky="nsew")

        # Display panel (terminal)
        self.display_frame = tk.Frame(self, bg=self.theme.bg)
        self.display_frame.grid(row=0, column=1, sticky="nsew")
        self.terminal = ctk.ScrolledText(...)

        # Status panel
        self.status_panel = StatusPanel(self, self.theme)
        self.status_panel.grid(row=1, column=1, sticky="nsew")

        # Wire callbacks
        self._wire_callbacks()
```

### Integrating EditorController into config_ui/App.py

The `EditorController` separates configuration logic from Configurator UI:

#### 1. Initialize controller in Configurator

```python
class Configurator:
    def __init__(self, master, theme, yml_path=None):
        # ... existing setup ...

        # Create controller instead of managing state directly
        self.controller = EditorController()

        # Register callbacks
        self.controller.register_callback('pipeline_created', self._on_pipeline_created)
        self.controller.register_callback('pipeline_loaded', self._on_pipeline_loaded)
        self.controller.register_callback('step_selected', self._on_step_selected)
        self.controller.register_callback('pipeline_changed', self._on_pipeline_changed)
        # ... more callbacks ...

        # Load or create pipeline
        if yml_path:
            success, error = self.controller.load_pipeline(Path(yml_path))
            if not success:
                messagebox.showerror("Error", error)
        else:
            self.controller.create_new_pipeline()
```

#### 2. Replace direct state manipulation

**Before:**
```python
def create_pipeline_step(self, step_type):
    self.STEP = step_type
    self.STEP_INDEX = len(self.CONFIG)
    self.CONFIG[self.STEP_INDEX] = {"step_general/step_type": step_type}
    # ... complex state management ...
```

**After:**
```python
def create_pipeline_step(self, step_type):
    step = self.controller.add_step(step_type)
    # Controller handles state and triggers callbacks
```

#### 3. Implement callbacks

```python
def _on_step_selected(self, index, step):
    """Handle step selection."""
    self.STEP_INDEX = index
    self.STEP = step.step_type
    self._update_editor()

def _on_pipeline_changed(self, pipeline):
    """Handle pipeline modifications."""
    self._update_pipeline()
    self.status_label.config(
        text="UNVALIDATED",
        bg=self.theme.yellow
    )
```

#### 4. Simplify validation

**Before:**
```python
def validate_full(self, return_config=False):
    # ... 40+ lines of validation logic ...
```

**After:**
```python
def validate_full(self):
    """Validate using controller."""
    is_valid, message = self.controller.validate_full()

    if is_valid:
        self.status_label.config(
            text="VALID",
            bg=self.theme.green
        )
    else:
        self.status_label.config(
            text="INVALID",
            bg=self.theme.red
        )

    messagebox.showinfo("Validation", message)
```

## Migration Strategy

### Recommended Approach: Incremental Migration

1. **Start with Controllers** (Low Risk)
   - Controllers can coexist with existing code
   - Add controller instances to existing classes
   - Gradually move logic from methods to controller
   - Test each migration independently

2. **Then UI Components** (Medium Risk)
   - Components can replace sections one at a time
   - Keep old code commented out during transition
   - Test each component replacement

3. **Finally, Complete Refactor** (Optional)
   - Once confident, create new main classes
   - Use components and controllers exclusively
   - Remove old code

### Testing Strategy

For each integrated component:

1. **Unit Tests**: Test controllers in isolation
   ```python
   def test_experiment_controller_start():
       controller = ExperimentController(test_config_path)
       assert controller.start_experiment(starting_slice=1)
       assert controller.state.is_running
   ```

2. **Integration Tests**: Test with UI
   - Manually test all workflows
   - Verify callbacks trigger correctly
   - Check error handling

3. **Regression Tests**: Ensure existing features work
   - Run through all experiment scenarios
   - Test edge cases (stops, errors, etc.)

## Benefits of Completion

Once fully integrated, you'll have:

1. **Testable Code**: Controllers can be unit tested without GUI
2. **Reusable Components**: UI widgets can be used in other contexts
3. **Clear Separation**: Easy to understand where logic lives
4. **Easier Debugging**: Smaller, focused classes
5. **Better Extensibility**: Easy to add features to right component

## Example: Complete Integration for One Feature

Here's how the "stop after slice" feature works with the new architecture:

**1. User clicks button → UI**
```python
# In ControlPanel
self.on_stop_slice()  # Callback to parent
```

**2. Parent triggers controller → Controller**
```python
# In MainApplication
def stop_slice(self):
    self.experiment_controller.request_stop_after_slice()
```

**3. Controller updates state → State**
```python
# In ExperimentController
def request_stop_after_slice(self):
    self.state.should_stop_slice = True
    self._notify('stop_requested', 'slice')
```

**4. Experiment loop checks state → Logic**
```python
# In ExperimentController._run_experiment_loop
if self.state.should_stop_slice:
    break
```

**5. Controller notifies UI → UI Update**
```python
# In MainApplication callback
def _on_experiment_stopped(self, final_slice, final_step):
    print("Stopped after slice")
    self._update_exp_control_buttons()
```

Clear, traceable flow through well-defined layers!

## Questions or Issues?

When integrating these changes, if you encounter:

- **Callback confusion**: Check the event registration in controller `__init__`
- **State synchronization**: Use controller as single source of truth
- **UI not updating**: Ensure callbacks call `self.update_idletasks()`
- **Testing difficulties**: Controllers should work without UI - check dependencies

## Next Steps

1. Review this guide
2. Choose integration approach (incremental recommended)
3. Start with one controller (ExperimentController recommended)
4. Test thoroughly
5. Gradually migrate remaining functionality
6. Remove old code once confident

Good luck! The architecture is now in place - it's just a matter of wiring it up.
