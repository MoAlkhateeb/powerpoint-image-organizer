from pathlib import Path
from typing import List

from PIL import Image
import customtkinter as ctk
from CTkToolTip import CTkToolTip
from tkinter import filedialog, messagebox, ttk

from pptx_generator import PPTXGenerator
from pptx_settings import PPTXSettings


class PPTXGeneratorGUI:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("PPTX Generator")
        self.root.geometry("950x750")

        self.pptx_generator = PPTXGenerator()
        self.settings = PPTXSettings()
        self.selected_images: List[Path] = []
        self.current_presentation_path = None

        self.create_widgets()

    def create_widgets(self):
        # Main frame
        main_frame = ctk.CTkFrame(self.root)
        main_frame.pack(fill=ctk.BOTH, expand=True, padx=20, pady=20)

        # Left panel for settings
        left_panel = ctk.CTkFrame(main_frame)
        left_panel.pack(side=ctk.LEFT, fill=ctk.BOTH, expand=True, padx=(0, 10))

        # Right panel for image selection and preview
        right_panel = ctk.CTkFrame(main_frame)
        right_panel.pack(side=ctk.RIGHT, fill=ctk.BOTH, expand=True, padx=(10, 0))

        # Settings widgets
        self.create_settings_widgets(left_panel)

        # Image selection and preview widgets
        self.create_image_widgets(right_panel)

        # Bottom panel for generate button
        bottom_panel = ctk.CTkFrame(self.root)
        bottom_panel.pack(side=ctk.BOTTOM, fill=ctk.X, padx=20, pady=20)

        self.create_action_buttons(bottom_panel)

    def create_settings_widgets(self, parent):
        ctk.CTkLabel(
            parent, text="Presentation Settings", font=("Arial", 16, "bold")
        ).pack(pady=(0, 10))

        # Presentation file selection
        file_frame = ctk.CTkFrame(parent)
        file_frame.pack(fill=ctk.X, pady=(0, 10))

        self.file_entry = ctk.CTkEntry(file_frame, width=200)
        self.file_entry.pack(side=ctk.LEFT, padx=(0, 5))
        self.create_tooltip(
            self.file_entry,
            "Enter the path to an existing presentation or a new file name",
        )

        file_btn = ctk.CTkButton(
            file_frame, text="Browse", command=self.select_presentation_file
        )
        file_btn.pack(side=ctk.LEFT)

        # Override/Append option
        self.override_var = ctk.BooleanVar(value=True)
        override_checkbox = ctk.CTkCheckBox(
            parent, text="Override existing presentation", variable=self.override_var
        )
        override_checkbox.pack(pady=(0, 10))
        self.create_tooltip(
            override_checkbox,
            "If checked, will override the existing presentation. If unchecked, will append to it.",
        )

        # Margins
        margins_frame = ctk.CTkFrame(parent)
        margins_frame.pack(fill=ctk.X, pady=(0, 10))

        ctk.CTkLabel(margins_frame, text="Margins (inches):").grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )
        self.top_margin = ctk.CTkEntry(margins_frame, width=50)
        self.top_margin.grid(row=0, column=1, padx=5, pady=5)
        self.top_margin.insert(0, str(self.settings.top_margin.inches))
        self.create_tooltip(self.top_margin, "Top margin in inches")

        self.left_margin = ctk.CTkEntry(margins_frame, width=50)
        self.left_margin.grid(row=0, column=2, padx=5, pady=5)
        self.left_margin.insert(0, str(self.settings.left_margin.inches))
        self.create_tooltip(self.left_margin, "Left margin in inches")

        self.right_margin = ctk.CTkEntry(margins_frame, width=50)
        self.right_margin.grid(row=1, column=1, padx=5, pady=5)
        self.right_margin.insert(0, str(self.settings.right_margin.inches))
        self.create_tooltip(self.right_margin, "Right margin in inches")

        self.bottom_margin = ctk.CTkEntry(margins_frame, width=50)
        self.bottom_margin.grid(row=1, column=2, padx=5, pady=5)
        self.bottom_margin.insert(0, str(self.settings.bottom_margin.inches))
        self.create_tooltip(self.bottom_margin, "Bottom margin in inches")

        # Center margins
        center_margins_frame = ctk.CTkFrame(parent)
        center_margins_frame.pack(fill=ctk.X, pady=(0, 10))

        ctk.CTkLabel(center_margins_frame, text="Center Margins (inches):").grid(
            row=0, column=0, sticky="w", padx=5, pady=5
        )
        self.h_center_margin = ctk.CTkEntry(center_margins_frame, width=50)
        self.h_center_margin.grid(row=0, column=1, padx=5, pady=5)
        self.h_center_margin.insert(0, str(self.settings.h_center_margin.inches))
        self.create_tooltip(self.h_center_margin, "Horizontal center margin in inches")

        self.v_center_margin = ctk.CTkEntry(center_margins_frame, width=50)
        self.v_center_margin.grid(row=0, column=2, padx=5, pady=5)
        self.v_center_margin.insert(0, str(self.settings.v_center_margin.inches))
        self.create_tooltip(self.v_center_margin, "Vertical center margin in inches")

        # Line width
        line_width_frame = ctk.CTkFrame(parent)
        line_width_frame.pack(fill=ctk.X, pady=(0, 10))

        ctk.CTkLabel(line_width_frame, text="Line Width (pt):").pack(
            side=ctk.LEFT, padx=5, pady=5
        )
        self.line_width = ctk.CTkEntry(line_width_frame, width=50)
        self.line_width.pack(side=ctk.LEFT, padx=5, pady=5)
        self.line_width.insert(0, str(self.settings.line_width.pt))
        self.create_tooltip(self.line_width, "Line width in points")

        # Color
        color_frame = ctk.CTkFrame(parent)
        color_frame.pack(fill=ctk.X, pady=(0, 10))

        ctk.CTkLabel(color_frame, text="Color:").pack(side=ctk.LEFT, padx=5, pady=5)
        self.color = ctk.CTkEntry(color_frame, width=100)
        self.color.pack(side=ctk.LEFT, padx=5, pady=5)
        self.color.insert(0, "#{:02x}{:02x}{:02x}".format(*self.settings.color))
        self.create_tooltip(self.color, "Color in hex format (e.g., #FF0000 for red)")

        # Rounded corners
        self.rounded = ctk.CTkCheckBox(parent, text="Rounded Corners")
        self.rounded.pack(pady=(0, 10))
        self.rounded.select() if self.settings.rounded else self.rounded.deselect()
        self.create_tooltip(self.rounded, "Enable rounded corners for images")

    def create_image_widgets(self, parent):
        ctk.CTkLabel(parent, text="Image Selection", font=("Arial", 16, "bold")).pack(
            pady=(0, 10)
        )

        select_btn = ctk.CTkButton(
            parent, text="Add Images", command=self.select_images
        )
        select_btn.pack(pady=(0, 10))

        self.image_listbox = ttk.Treeview(
            parent, columns=("Order", "Filename"), show="headings", height=10
        )
        self.image_listbox.heading("Order", text="Order")
        self.image_listbox.heading("Filename", text="Filename")
        self.image_listbox.column("Order", width=50)
        self.image_listbox.column("Filename", width=200)
        self.image_listbox.pack(fill=ctk.BOTH, expand=True, pady=(0, 10))

        self.image_listbox.bind("<<TreeviewSelect>>", self.on_image_select)

        button_frame = ctk.CTkFrame(parent)
        button_frame.pack(fill=ctk.X, pady=(0, 10))

        move_up_btn = ctk.CTkButton(
            button_frame, text="Move Up", command=self.move_image_up
        )
        move_up_btn.pack(side=ctk.LEFT, padx=(0, 5))

        move_down_btn = ctk.CTkButton(
            button_frame, text="Move Down", command=self.move_image_down
        )
        move_down_btn.pack(side=ctk.LEFT, padx=(0, 5))

        delete_btn = ctk.CTkButton(
            button_frame, text="Delete", command=self.delete_selected_image
        )
        delete_btn.pack(side=ctk.LEFT, padx=(0, 5))

        delete_all_btn = ctk.CTkButton(
            button_frame, text="Delete All", command=self.delete_all_images
        )
        delete_all_btn.pack(side=ctk.LEFT)

        preview_frame = ctk.CTkFrame(parent)
        preview_frame.pack(fill=ctk.BOTH, expand=True)

        self.preview_label = ctk.CTkLabel(preview_frame, text="")
        self.preview_label.pack()

    def create_action_buttons(self, parent):
        generate_btn = ctk.CTkButton(
            parent, text="Generate Presentation", command=self.generate_presentation
        )
        generate_btn.pack(side=ctk.TOP, pady=10)

    def select_presentation_file(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pptx",
            filetypes=[("PowerPoint Presentation", "*.pptx")],
            title="Select Existing Presentation or Enter New Filename",
        )
        if file_path:
            self.file_entry.delete(0, ctk.END)
            self.file_entry.insert(0, file_path)
            self.current_presentation_path = file_path

    def select_images(self):
        filetypes = (
            ("Image files", "*.png *.jpg *.jpeg *.gif *.bmp"),
            ("All files", "*.*"),
        )
        image_paths = filedialog.askopenfilenames(filetypes=filetypes)
        new_images = [Path(path) for path in image_paths]
        self.selected_images.extend(new_images)
        self.update_image_listbox()

    def update_image_listbox(self):
        self.image_listbox.delete(*self.image_listbox.get_children())
        for index, image in enumerate(self.selected_images, start=1):
            self.image_listbox.insert("", "end", values=(index, image.name))

    def on_image_select(self, event):
        selected_item = self.image_listbox.selection()
        if selected_item:
            index = int(self.image_listbox.item(selected_item)["values"][0]) - 1
            self.show_image_preview(self.selected_images[index])

    def show_image_preview(self, image_path):
        image = Image.open(image_path)
        photo = ctk.CTkImage(image, size=(200, 200))
        self.preview_label.configure(image=photo)
        self.preview_label.image = photo

    def move_image_up(self):
        selected_item = self.image_listbox.selection()
        if selected_item:
            index = self.image_listbox.index(selected_item)
            if index > 0:
                self.selected_images[index], self.selected_images[index - 1] = (
                    self.selected_images[index - 1],
                    self.selected_images[index],
                )
                self.update_image_listbox()
                self.image_listbox.selection_set(
                    self.image_listbox.get_children()[index - 1]
                )

    def move_image_down(self):
        selected_item = self.image_listbox.selection()
        if selected_item:
            index = self.image_listbox.index(selected_item)
            if index < len(self.selected_images) - 1:
                self.selected_images[index], self.selected_images[index + 1] = (
                    self.selected_images[index + 1],
                    self.selected_images[index],
                )
                self.update_image_listbox()
                self.image_listbox.selection_set(
                    self.image_listbox.get_children()[index + 1]
                )

    def delete_selected_image(self):
        selected_items = self.image_listbox.selection()
        if selected_items:
            for item in reversed(selected_items):
                index = self.image_listbox.index(item)
                del self.selected_images[index]
            self.update_image_listbox()
            self.reset_image_preview()

    def delete_all_images(self):
        self.selected_images.clear()
        self.update_image_listbox()
        self.reset_image_preview()

    def reset_image_preview(self):
        self.preview_label.configure(image=None)
        self.preview_label.image = None

    def generate_presentation(self):
        try:
            self.update_settings()
            presentation_path = self.file_entry.get() or "new_presentation.pptx"
            self.pptx_generator.create_presentation(
                presentation_path, override=self.override_var.get()
            )
            self.pptx_generator.add_images(self.selected_images)
            self.pptx_generator.save_presentation(presentation_path)
            messagebox.showinfo(
                "Success", f"Presentation generated and saved to {presentation_path}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate presentation: {str(e)}")

    def update_settings(self):
        self.settings.top_margin = float(self.top_margin.get())
        self.settings.left_margin = float(self.left_margin.get())
        self.settings.right_margin = float(self.right_margin.get())
        self.settings.bottom_margin = float(self.bottom_margin.get())
        self.settings.h_center_margin = float(self.h_center_margin.get())
        self.settings.v_center_margin = float(self.v_center_margin.get())
        self.settings.line_width = float(self.line_width.get())
        self.settings.color = self.color.get()
        self.settings.rounded = self.rounded.get()

        self.pptx_generator.settings = self.settings

    def create_tooltip(self, widget, text):
        CTkToolTip(widget, message=text)

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = PPTXGeneratorGUI()
    app.run()
