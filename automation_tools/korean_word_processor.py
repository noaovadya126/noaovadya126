import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import openai
import os
from pathlib import Path

class KoreanWordProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Korean Word Processor")
        self.root.geometry("700x500")
        self.root.configure(bg='#f0f0f0')
        
        self.file_path = ""
        self.df = None
        self.processed_df = None
        
        self.setup_ui()
        
    def setup_ui(self):
        title_label = tk.Label(self.root, text="Korean Word Processor", font=("Arial", 16, "bold"), bg='#f0f0f0')
        title_label.pack(pady=20)
        
        file_frame = tk.Frame(self.root, bg='#f0f0f0')
        file_frame.pack(pady=10, padx=20, fill='x')
        
        tk.Label(file_frame, text="Excel File:", font=("Arial", 10), bg='#f0f0f0').pack(anchor='w')
        
        file_select_frame = tk.Frame(file_frame, bg='#f0f0f0')
        file_select_frame.pack(fill='x', pady=5)
        
        self.file_entry = tk.Entry(file_select_frame, font=("Arial", 10), width=50)
        self.file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        browse_btn = tk.Button(file_select_frame, text="Browse", command=self.browse_file, 
                              bg='#4CAF50', fg='white', font=("Arial", 10), relief='flat')
        browse_btn.pack(side='right')
        
        button_frame = tk.Frame(self.root, bg='#f0f0f0')
        button_frame.pack(pady=20)
        
        self.process_btn = tk.Button(button_frame, text="Process", command=self.process_file, 
                                    bg='#2196F3', fg='white', font=("Arial", 12, "bold"), 
                                    relief='flat', width=15, height=2)
        self.process_btn.pack(side='left', padx=10)
        
        self.test_btn = tk.Button(button_frame, text="Test First 50", command=self.test_first_50, 
                                 bg='#9C27B0', fg='white', font=("Arial", 12, "bold"), 
                                 relief='flat', width=15, height=2)
        self.test_btn.pack(side='left', padx=10)
        
        self.save_btn = tk.Button(button_frame, text="Save As", command=self.save_file, 
                                 bg='#FF9800', fg='white', font=("Arial", 12, "bold"), 
                                 relief='flat', width=15, height=2, state='disabled')
        self.save_btn.pack(side='left', padx=10)
        
        self.progress = ttk.Progressbar(self.root, mode='determinate')
        self.progress.pack(pady=20, padx=20, fill='x')
        
        self.status_label = tk.Label(self.root, text="Ready", font=("Arial", 10), bg='#f0f0f0')
        self.status_label.pack(pady=10)
        
        preview_frame = tk.Frame(self.root, bg='#f0f0f0')
        preview_frame.pack(pady=10, padx=20, fill='both', expand=True)
        
        tk.Label(preview_frame, text="Preview:", font=("Arial", 10, "bold"), bg='#f0f0f0').pack(anchor='w')
        
        self.preview_text = tk.Text(preview_frame, height=8, width=80, font=("Consolas", 9))
        self.preview_text.pack(fill='both', expand=True)
        
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.status_label.config(text="File selected: " + os.path.basename(file_path))
            
    def test_first_50(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first!")
            return
            
        try:
            self.df = pd.read_excel(self.file_path)
            if '어휘' not in self.df.columns:
                messagebox.showerror("Error", "Column '어휘' not found in the Excel file!")
                return
                
            self.processed_df = self.df.copy()
            self.processed_df['AI_Output'] = ""
            
            test_rows = min(50, len(self.df))
            self.progress['maximum'] = test_rows
            self.progress['value'] = 0
            
            self.status_label.config(text=f"Testing first {test_rows} words... Please wait.")
            self.root.update()
            
            preview_data = []
            
            for index in range(test_rows):
                row = self.df.iloc[index]
                korean_word = str(row['어휘'])
                pos = str(row['품사']) if '품사' in self.df.columns else ""
                original_guide = str(row['길잡이 말']) if '길잡이 말' in self.df.columns else ""
                
                ai_output = self.get_ai_assistance(korean_word, pos)
                self.processed_df.at[index, 'AI_Output'] = ai_output
                
                preview_data.append(f"{index+1:2d}. {korean_word:10s} ({pos:5s}) | Original: {original_guide:20s} | AI: {ai_output}")
                
                self.progress['value'] = index + 1
                self.root.update()
                
            self.preview_text.delete(1.0, tk.END)
            self.preview_text.insert(1.0, "\n".join(preview_data))
            
            self.status_label.config(text=f"Test completed! Processed {test_rows} words.")
            self.save_btn.config(state='normal')
            messagebox.showinfo("Success", f"Test completed successfully! Processed {test_rows} words.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing file: {str(e)}")
            self.status_label.config(text="Error occurred")
            
    def process_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "Please select a file first!")
            return
            
        try:
            self.df = pd.read_excel(self.file_path)
            if '어휘' not in self.df.columns:
                messagebox.showerror("Error", "Column '어휘' not found in the Excel file!")
                return
                
            self.processed_df = self.df.copy()
            self.processed_df['AI_Output'] = ""
            
            self.progress['maximum'] = len(self.df)
            self.progress['value'] = 0
            
            self.status_label.config(text="Processing... Please wait.")
            self.root.update()
            
            for index, row in self.df.iterrows():
                korean_word = str(row['어휘'])
                pos = str(row['품사']) if '품사' in self.df.columns else ""
                
                ai_output = self.get_ai_assistance(korean_word, pos)
                self.processed_df.at[index, 'AI_Output'] = ai_output
                
                self.progress['value'] = index + 1
                self.root.update()
                
            self.status_label.config(text="Processing completed!")
            self.save_btn.config(state='normal')
            messagebox.showinfo("Success", "File processed successfully!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing file: {str(e)}")
            self.status_label.config(text="Error occurred")
            
    def get_ai_assistance(self, korean_word, pos):
        try:
            openai.api_key = "your-api-key-here"
            
            prompt = f"Given the Korean word '{korean_word}' with part of speech '{pos}', provide a natural Korean phrase or sentence that uses this word correctly. Return only the Korean text, no explanations."
            
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are a Korean language expert. Provide natural Korean phrases using the given word."},
                    {"role": "user", "content": prompt}
                ],
                max_tokens=50,
                temperature=0.7
            )
            
            return response.choices[0].message.content.strip()
            
        except Exception as e:
            return f"Error: {str(e)}"
            
    def save_file(self):
        if self.processed_df is None:
            messagebox.showerror("Error", "No processed data to save!")
            return
            
        save_path = filedialog.asksaveasfilename(
            title="Save Processed File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if save_path:
            try:
                self.processed_df.to_excel(save_path, index=False)
                messagebox.showinfo("Success", f"File saved successfully to:\n{save_path}")
                self.status_label.config(text="File saved successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Error saving file: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = KoreanWordProcessor(root)
    root.mainloop()
