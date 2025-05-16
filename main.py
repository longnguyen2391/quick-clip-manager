import keyboard
import tkinter as tk
from tkinter import ttk, messagebox
import pyperclip
import sqlite3
from datetime import datetime
import sys
import threading
import traceback
from infi.systray import SysTrayIcon

class QuickClipManager:
    def __init__(self):
        # Initialize state variables first
        self.is_quick_window_visible = False
        self.quick_window = None
        self.suggestion_window = None
        self.root = None
        self.suggestion_listbox = None
        self.current_suggestions = []
        self.running = True
        self.tray_icon = None
        
        # Then setup the application
        self.setup_database()
        self.setup_main_window()
        self.setup_hotkey()
        self.setup_tray()

    def setup_tray(self):
        """Setup system tray icon and menu"""
        try:
            menu_options = (
                ("Show Window", None, self.show_window),
                ("Exit", None, self.quit_application)
            )
            
            self.tray_icon = SysTrayIcon(
                None,  # Use default icon
                "Quick Clip Manager",
                menu_options,
                on_quit=self.quit_application
            )
            
            # Start the tray icon
            self.tray_icon.start()
            
        except Exception as e:
            print(f"Error setting up system tray: {str(e)}")
            print(traceback.format_exc())

    def hide_to_tray(self):
        """Hide the main window to system tray"""
        try:
            self.root.withdraw()
            self.show_notification("Running in background")
        except Exception as e:
            print(f"Error hiding to tray: {str(e)}")

    def show_window(self, systray=None):
        """Show the main window from system tray"""
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception as e:
            print(f"Error showing window: {str(e)}")

    def setup_database(self):
        """Initialize SQLite database and create tables if they don't exist"""
        try:
            self.conn = sqlite3.connect('quickclip.db', check_same_thread=False)
            self.cursor = self.conn.cursor()
            
            # Create aliases table
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS aliases (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    alias TEXT UNIQUE NOT NULL,
                    content TEXT NOT NULL,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            self.conn.commit()
        except Exception as e:
            messagebox.showerror("Database Error", f"Failed to setup database: {str(e)}")
            raise

    def setup_main_window(self):
        """Setup the main management window"""
        self.root = tk.Tk()
        self.root.title("Quick Clip Manager")
        self.root.geometry("600x400")
        
        # Set proper window closing behavior
        self.root.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

        # Create main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Alias input
        ttk.Label(main_frame, text="Alias:").grid(row=0, column=0, sticky=tk.W)
        self.alias_entry = ttk.Entry(main_frame, width=20)
        self.alias_entry.grid(row=0, column=1, sticky=(tk.W, tk.E))

        # Content input
        ttk.Label(main_frame, text="Content:").grid(row=1, column=0, sticky=tk.W)
        self.content_text = tk.Text(main_frame, width=40, height=4)
        self.content_text.grid(row=1, column=1, sticky=(tk.W, tk.E))

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Add/Update", command=self.save_alias).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Delete", command=self.delete_alias).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Hide", command=self.hide_to_tray).pack(side=tk.LEFT, padx=5)
        
        # Exit button that forces immediate exit
        exit_btn = ttk.Button(button_frame, text="Exit")
        exit_btn.pack(side=tk.LEFT, padx=5)
        exit_btn.configure(command=lambda: (
            threading.Thread(target=self.force_exit, daemon=True).start()
        ))

        # Treeview for aliases
        self.tree = ttk.Treeview(main_frame, columns=('Alias', 'Content'), show='headings')
        self.tree.heading('Alias', text='Alias')
        self.tree.heading('Content', text='Content')
        self.tree.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=3, column=2, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.bind('<<TreeviewSelect>>', self.on_select)
        self.load_aliases()

        # Add hotkey status label
        self.status_label = ttk.Label(main_frame, text="Running in background (Ctrl+F1 to toggle quick input)")
        self.status_label.grid(row=4, column=0, columnspan=2, pady=5)

        # Create quick input window last
        self.create_quick_input_window()

    def setup_hotkey(self):
        """Setup global hotkey in a separate thread"""
        def hotkey_thread():
            while self.running:
                try:
                    keyboard.wait('ctrl+f1')
                    if self.running:  # Check if still running
                        self.root.after(0, self.toggle_quick_input)
                except Exception as e:
                    print(f"Hotkey error: {str(e)}")
                    if not self.running:
                        break

        self.hotkey_thread = threading.Thread(target=hotkey_thread, daemon=True)
        self.hotkey_thread.start()

    def quit_application(self, systray=None):
        """Properly close the application"""
        print("Starting application shutdown...")
        
        # Force immediate exit without trying to clean up tkinter windows
        try:
            # Stop hotkey thread
            self.running = False
            
            # Unhook keyboard
            try:
                keyboard.unhook_all()
            except:
                pass

            # Stop tray icon
            try:
                if self.tray_icon:
                    self.tray_icon.shutdown()
            except:
                pass

            # Close database
            try:
                if hasattr(self, 'conn'):
                    self.conn.close()
            except:
                pass

        except Exception as e:
            print(f"Error during shutdown: {str(e)}")
        
        print("Forcing immediate exit...")
        # Force exit without attempting to close windows
        import os
        os._exit(0)

    def minimize_to_tray(self):
        """Minimize the window instead of closing it"""
        self.root.iconify()

    def create_quick_input_window(self):
        """Create the quick input window"""
        if self.quick_window is not None:
            self.quick_window.destroy()
            
        self.quick_window = tk.Toplevel(self.root)
        self.quick_window.title("Quick Input")
        self.quick_window.attributes('-topmost', True)
        
        # Remove window decorations but keep it in taskbar
        self.quick_window.overrideredirect(True)
        
        # Create a frame with a border
        frame = ttk.Frame(self.quick_window, relief='solid', borderwidth=1)
        frame.pack(fill='both', expand=True)
        
        self.quick_entry = ttk.Entry(frame, width=30)
        self.quick_entry.pack(padx=5, pady=5)
        
        # Create suggestion window
        self.create_suggestion_window()
        
        # Bind events
        self.quick_entry.bind('<Return>', self.process_quick_input)
        self.quick_entry.bind('<Escape>', lambda e: self.toggle_quick_input())
        self.quick_entry.bind('<KeyRelease>', self.update_suggestions)
        self.quick_entry.bind('<Down>', self.focus_suggestions)
        self.quick_window.protocol("WM_DELETE_WINDOW", lambda: self.toggle_quick_input())
        
        # Initially hide the window
        self.quick_window.withdraw()
        self.is_quick_window_visible = False

    def create_suggestion_window(self):
        """Create the suggestion dropdown window"""
        if self.suggestion_window is not None:
            self.suggestion_window.destroy()

        self.suggestion_window = tk.Toplevel(self.quick_window)
        self.suggestion_window.withdraw()
        self.suggestion_window.overrideredirect(True)
        self.suggestion_window.attributes('-topmost', True)

        # Create frame with border
        frame = ttk.Frame(self.suggestion_window, relief='solid', borderwidth=1)
        frame.pack(fill='both', expand=True)

        # Create listbox for suggestions
        self.suggestion_listbox = tk.Listbox(
            frame,
            width=40,
            height=5,
            font=('TkDefaultFont', 9),
            selectmode=tk.SINGLE,
            activestyle='dotbox'
        )
        self.suggestion_listbox.pack(fill='both', expand=True)

        # Bind events
        self.suggestion_listbox.bind('<Return>', self.use_suggestion)
        self.suggestion_listbox.bind('<Escape>', lambda e: self.hide_suggestions())
        self.suggestion_listbox.bind('<Double-Button-1>', self.use_suggestion)
        self.suggestion_listbox.bind('<Button-1>', self.handle_suggestion_click)
        self.suggestion_listbox.bind('<Up>', self.handle_suggestion_keys)
        self.suggestion_listbox.bind('<Down>', self.handle_suggestion_keys)
        self.suggestion_listbox.bind('<Tab>', self.handle_suggestion_keys)

        # Bind quick entry events for navigation
        self.quick_entry.bind('<Up>', self.handle_quick_entry_keys)
        self.quick_entry.bind('<Down>', self.handle_quick_entry_keys)
        self.quick_entry.bind('<Tab>', self.handle_quick_entry_keys)

    def update_suggestions(self, event=None):
        """Update suggestion list based on current input"""
        try:
            text = self.quick_entry.get().strip()
            
            if text.startswith('/'):
                search_text = text[1:].lower()
                if search_text:
                    # Search for matching aliases using partial match
                    self.cursor.execute('''
                        SELECT alias, content FROM aliases 
                        WHERE LOWER(alias) LIKE ? 
                        ORDER BY 
                            CASE 
                                WHEN LOWER(alias) = ? THEN 0
                                WHEN LOWER(alias) LIKE ? THEN 1
                                ELSE 2
                            END,
                            LENGTH(alias),
                            alias
                        LIMIT 10
                    ''', (f'%{search_text}%', search_text, f'{search_text}%'))
                    
                    matches = self.cursor.fetchall()
                    print(f"Found matches for '{search_text}': {matches}")  # Debug print
                    
                    if matches:
                        self.show_suggestions(matches)
                    else:
                        self.hide_suggestions()
                else:
                    # If just '/' is typed, show all aliases
                    self.cursor.execute('''
                        SELECT alias, content FROM aliases 
                        ORDER BY alias LIMIT 10
                    ''')
                    matches = self.cursor.fetchall()
                    if matches:
                        self.show_suggestions(matches)
                    else:
                        self.hide_suggestions()
            else:
                self.hide_suggestions()
                
        except Exception as e:
            print(f"Error updating suggestions: {str(e)}")
            print(traceback.format_exc())

    def show_suggestions(self, matches):
        """Show suggestion window with matches"""
        try:
            self.suggestion_listbox.delete(0, tk.END)
            self.current_suggestions = []
            
            max_alias_length = max(len(alias) for alias, _ in matches)
            
            for alias, content in matches:
                # Format the display text with aligned columns
                display_text = f"{alias.ljust(max_alias_length + 2)} | {content[:50]}{'...' if len(content) > 50 else ''}"
                self.suggestion_listbox.insert(tk.END, display_text)
                self.current_suggestions.append((alias, content))

            # Adjust listbox width based on content
            max_width = max(len(text) for text in self.suggestion_listbox.get(0, tk.END))
            self.suggestion_listbox.configure(width=max_width)

            # Position suggestion window below quick input
            x = self.quick_window.winfo_x()
            y = self.quick_window.winfo_y() + self.quick_window.winfo_height()
            
            # Calculate required height based on number of items
            item_height = self.suggestion_listbox.winfo_reqheight() // max(len(matches), 1)
            height = min(len(matches) * item_height + 4, 200)  # Max height of 200 pixels
            
            self.suggestion_window.geometry(f"+{x}+{y}")
            self.suggestion_listbox.configure(height=min(len(matches), 10))
            
            self.suggestion_window.deiconify()
            self.suggestion_window.lift()
            
        except Exception as e:
            print(f"Error showing suggestions: {str(e)}")
            print(traceback.format_exc())

    def hide_suggestions(self):
        """Hide the suggestion window"""
        try:
            if self.suggestion_window:
                self.suggestion_window.withdraw()
            self.current_suggestions = []
        except Exception as e:
            print(f"Error hiding suggestions: {str(e)}")

    def focus_suggestions(self, event):
        """Move focus to suggestion list"""
        if self.suggestion_window.winfo_viewable():
            self.suggestion_listbox.focus_set()
            if self.suggestion_listbox.size() > 0:
                self.suggestion_listbox.selection_set(0)
            return "break"

    def handle_quick_entry_keys(self, event):
        """Handle keyboard navigation from quick entry"""
        if not self.suggestion_window.winfo_viewable():
            return

        if event.keysym in ('Down', 'Tab'):
            self.suggestion_listbox.focus_set()
            self.suggestion_listbox.selection_clear(0, tk.END)
            self.suggestion_listbox.selection_set(0)
            self.suggestion_listbox.activate(0)
            return "break"
        return None

    def handle_suggestion_keys(self, event):
        """Handle keyboard navigation in suggestion list"""
        current_selection = self.suggestion_listbox.curselection()
        size = self.suggestion_listbox.size()
        
        if not current_selection:
            if event.keysym in ('Down', 'Tab'):
                self.suggestion_listbox.selection_set(0)
                self.suggestion_listbox.activate(0)
                return "break"
            elif event.keysym == 'Up':
                self.quick_entry.focus_set()
                return "break"
        else:
            current_index = current_selection[0]
            
            if event.keysym == 'Up':
                if current_index == 0:
                    self.suggestion_listbox.selection_clear(0)
                    self.quick_entry.focus_set()
                    return "break"
                else:
                    self.suggestion_listbox.selection_clear(0, tk.END)
                    self.suggestion_listbox.selection_set(current_index - 1)
                    self.suggestion_listbox.activate(current_index - 1)
                    self.suggestion_listbox.see(current_index - 1)
                    return "break"
            elif event.keysym in ('Down', 'Tab'):
                if current_index < size - 1:
                    self.suggestion_listbox.selection_clear(0, tk.END)
                    self.suggestion_listbox.selection_set(current_index + 1)
                    self.suggestion_listbox.activate(current_index + 1)
                    self.suggestion_listbox.see(current_index + 1)
                    return "break"
        return None

    def handle_suggestion_click(self, event):
        """Handle mouse click on suggestion list"""
        if self.suggestion_listbox.curselection():
            self.suggestion_listbox.focus_set()
        return None

    def use_suggestion(self, event):
        """Use the selected suggestion"""
        try:
            if self.suggestion_listbox.curselection():
                index = self.suggestion_listbox.curselection()[0]
                if 0 <= index < len(self.current_suggestions):
                    alias, content = self.current_suggestions[index]
                    print(f"Using suggestion: {alias} with content: {content}")  # Debug print
                    pyperclip.copy(content)
                    self.hide_quick_input()
                    self.show_notification(f"Copied: {alias}")
        except Exception as e:
            print(f"Error using suggestion: {str(e)}")
            print(traceback.format_exc())

    def show_quick_input(self):
        """Show the quick input window at the center of the screen"""
        try:
            if not self.quick_window:
                self.create_quick_input_window()
                
            # Get screen dimensions
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            window_width = 250
            window_height = 35
            
            # Calculate position
            x = (screen_width - window_width) // 2
            y = (screen_height - window_height) // 2
            
            # Configure window
            self.quick_window.geometry(f'{window_width}x{window_height}+{x}+{y}')
            self.quick_entry.delete(0, tk.END)
            
            # Show window
            self.quick_window.deiconify()
            self.quick_window.lift()
            self.quick_window.focus_force()
            self.quick_entry.focus_set()
            
            self.is_quick_window_visible = True
            
        except Exception as e:
            print(f"Error showing quick input: {str(e)}")
            print(traceback.format_exc())

    def hide_quick_input(self):
        """Hide the quick input window"""
        try:
            if self.quick_window:
                self.quick_window.withdraw()
            if self.suggestion_window:
                self.suggestion_window.withdraw()
            self.is_quick_window_visible = False
        except Exception as e:
            print(f"Error hiding quick input: {str(e)}")
            print(traceback.format_exc())

    def toggle_quick_input(self):
        """Toggle the quick input window visibility"""
        try:
            if self.is_quick_window_visible:
                self.hide_quick_input()
            else:
                self.show_quick_input()
        except Exception as e:
            print(f"Error toggling quick input: {str(e)}")
            print(traceback.format_exc())

    def process_quick_input(self, event):
        """Process the quick input when Enter is pressed"""
        try:
            text = self.quick_entry.get().strip()
            if text.startswith('/'):
                search_text = text[1:].lower()
                print(f"Processing input: {search_text}")  # Debug print
                
                try:
                    # First try exact match
                    self.cursor.execute('''
                        SELECT content FROM aliases 
                        WHERE LOWER(alias) = ? COLLATE NOCASE
                    ''', (search_text,))
                    result = self.cursor.fetchone()
                    
                    if not result:
                        # If no exact match, try partial match
                        self.cursor.execute('''
                            SELECT content FROM aliases 
                            WHERE LOWER(alias) LIKE ? COLLATE NOCASE
                            ORDER BY LENGTH(alias), alias LIMIT 1
                        ''', (f'%{search_text}%',))
                        result = self.cursor.fetchone()
                    
                    print(f"Search result: {result}")  # Debug print
                    
                    if result:
                        content = result[0]
                        print(f"Copying content: {content}")  # Debug print
                        
                        # Try copying with error handling
                        try:
                            pyperclip.copy(content)
                            print("Content copied successfully")
                            self.hide_quick_input()
                            self.show_notification(f"Copied content to clipboard")
                        except Exception as e:
                            print(f"Clipboard error: {str(e)}")
                            self.show_notification(f"Error copying to clipboard: {str(e)}")
                    else:
                        print(f"No match found for: {search_text}")
                        self.show_notification(f"No match found for: {search_text}")
                        
                except sqlite3.Error as e:
                    print(f"Database error: {str(e)}")
                    self.show_notification(f"Database error: {str(e)}")
            
            self.quick_entry.delete(0, tk.END)
            
        except Exception as e:
            print(f"Error in process_quick_input: {str(e)}")
            print(traceback.format_exc())
            self.show_notification(f"Error: {str(e)}")

    def show_notification(self, message):
        """Show a temporary notification"""
        try:
            notification = tk.Toplevel(self.root)
            notification.attributes('-topmost', True)
            notification.overrideredirect(True)
            
            # Position near the quick input window
            x = self.quick_window.winfo_x()
            y = self.quick_window.winfo_y() + self.quick_window.winfo_height() + 5
            
            notification.geometry(f"+{x}+{y}")
            
            # Create a frame with a border
            frame = ttk.Frame(notification, relief='solid', borderwidth=1)
            frame.pack(fill='both', expand=True)
            
            label = ttk.Label(frame, text=message, padding=5)
            label.pack()
            
            # Auto-close after 2 seconds
            self.root.after(2000, notification.destroy)
        except Exception as e:
            print(f"Error showing notification: {str(e)}")

    def save_alias(self):
        """Save or update an alias"""
        try:
            alias = self.alias_entry.get().strip()
            content = self.content_text.get("1.0", tk.END).strip()
            
            if not alias or not content:
                messagebox.showwarning("Warning", "Both alias and content are required!")
                return

            print(f"Saving alias: {alias} with content: {content}")
            
            self.cursor.execute('''
                INSERT INTO aliases (alias, content) VALUES (?, ?)
                ON CONFLICT(alias) DO UPDATE SET 
                    content = excluded.content,
                    updated_at = CURRENT_TIMESTAMP
            ''', (alias, content))
            self.conn.commit()
            print("Alias saved successfully")
            
            self.load_aliases()
            self.clear_inputs()
            messagebox.showinfo("Success", "Alias saved successfully!")
        except sqlite3.Error as e:
            print(f"Database error while saving: {str(e)}")
            messagebox.showerror("Error", f"Failed to save alias: {str(e)}")
        except Exception as e:
            print(f"Error while saving: {str(e)}")
            print(traceback.format_exc())
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")

    def delete_alias(self):
        """Delete selected alias"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an alias to delete!")
            return
        
        alias = self.tree.item(selection[0])['values'][0]
        if messagebox.askyesno("Confirm", f"Delete alias '{alias}'?"):
            self.cursor.execute('DELETE FROM aliases WHERE alias = ?', (alias,))
            self.conn.commit()
            self.load_aliases()
            self.clear_inputs()

    def load_aliases(self):
        """Load aliases into the treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        self.cursor.execute('SELECT alias, content FROM aliases ORDER BY alias')
        for row in self.cursor.fetchall():
            self.tree.insert('', tk.END, values=row)

    def on_select(self, event):
        """Handle treeview selection"""
        selection = self.tree.selection()
        if selection:
            alias, content = self.tree.item(selection[0])['values']
            self.alias_entry.delete(0, tk.END)
            self.alias_entry.insert(0, alias)
            self.content_text.delete("1.0", tk.END)
            self.content_text.insert("1.0", content)

    def clear_inputs(self):
        """Clear input fields"""
        self.alias_entry.delete(0, tk.END)
        self.content_text.delete("1.0", tk.END)

    def force_exit(self):
        """Force exit the application from a separate thread"""
        print("Force exit initiated...")
        import os
        os._exit(0)

    def run(self):
        """Start the application"""
        try:
            # Show initial notification
            self.show_notification("Quick Clip Manager is running")
            self.root.mainloop()
        except Exception as e:
            print(f"Error in main loop: {str(e)}")
        finally:
            # Force immediate exit
            self.force_exit()

if __name__ == "__main__":
    app = QuickClipManager()
    app.run() 