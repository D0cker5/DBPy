import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, simpledialog, messagebox, Toplevel
import pandas as pd
from fuzzywuzzy import fuzz
from plexapi.server import PlexServer
from plexapi.playlist import Playlist

# Plex connection details
PLEX_BASEURL = 'http://192.168.68.129:32400'
PLEX_TOKEN = '64EwJMdQg95LjMUKcKoE'
plex = PlexServer(PLEX_BASEURL, PLEX_TOKEN)

class PlaylistApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Plex Playlist Matcher")

        # Remove full-screen mode, allow standard window operations
        self.root.geometry("1200x800")  # Set an initial window size
        self.root.minsize(800, 600)  # Set minimum window size
        self.root.resizable(True, True)  # Allow window to be resizable

        # Variables
        self.csv_data = None
        self.results = []
        self.track_map = {}
        self.db_path = None

        # Create GUI Elements
        self.load_db_button = tk.Button(root, text="Select Database File", command=self.load_database)
        self.load_db_button.grid(row=0, column=0, padx=10, pady=10)

        self.load_button = tk.Button(root, text="Select CSV File", command=self.load_csv)
        self.load_button.grid(row=0, column=1, padx=10, pady=10)

        self.go_button = tk.Button(root, text="GO", command=self.process_tracks, state=tk.DISABLED)
        self.go_button.grid(row=0, column=2, padx=10, pady=10)

        self.progress_label = tk.Label(root, text="Progress: 0/0")
        self.progress_label.grid(row=1, column=0, columnspan=3)

        # Add a frame to contain TreeView and scrollbar
        frame = tk.Frame(root)
        frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")

        # Make rows/columns expandable
        root.grid_rowconfigure(2, weight=1)
        root.grid_columnconfigure(0, weight=1)

        # TreeView to display results, which resizes dynamically
        self.results_frame = ttk.Treeview(frame, columns=("Source Track", "Source Artist", "Matched Track", "Matched Artist", "Matched Album", "Score", "RatingKey"), show="headings")
        self.results_frame.heading("Source Track", text="Source Track")
        self.results_frame.heading("Source Artist", text="Source Artist")
        self.results_frame.heading("Matched Track", text="Matched Track")
        self.results_frame.heading("Matched Artist", text="Matched Artist")
        self.results_frame.heading("Matched Album", text="Matched Album")
        self.results_frame.heading("Score", text="Match Score")
        self.results_frame.heading("RatingKey", text="RatingKey")

        # Add vertical scrollbar
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.results_frame.yview)
        self.results_frame.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.results_frame.pack(fill="both", expand=True)

        # Configure tag for amber background
        self.results_frame.tag_configure('amber', background='#FFC04C')

        self.save_button = tk.Button(root, text="Save Playlist", command=self.save_playlist, state=tk.DISABLED)
        self.save_button.grid(row=3, column=0, columnspan=3, pady=10)

        # Context Menu for Right-click
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Remove", command=self.remove_selected_row)
        self.context_menu.add_command(label="Fix", command=self.fix_selected_row)
        self.context_menu.add_command(label="Fuzzy", command=self.fuzzy_selected_row)  # Add the 'Fuzzy' option

        # Bind a method to handle clicks in the TreeView
        self.results_frame.bind("<ButtonRelease-1>", self.on_row_click)
        self.results_frame.bind("<Button-3>", self.show_context_menu)







    def fuzzy_selected_row(self):
        """Fuzzy match the selected track against the database."""
        selected_item = self.results_frame.selection()[0]
        track_info = self.track_map.get(selected_item)
    
        if track_info:
            track = track_info["track"] if isinstance(track_info, dict) else track_info[0]
            artist = track_info["artist"] if isinstance(track_info, dict) else track_info[1]
            # Perform fuzzy matching and show results for manual selection
            self.perform_fuzzy_match(track, artist, selected_item)

    def perform_fuzzy_match(self, track, artist, item_id):
        """Re-query the database, fuzzy match track name and artist, and return the top 10 matches."""
        # Create a popup window to display results
        manual_window = Toplevel(self.root)
        manual_window.title(f"Fuzzy Match for {track} by {artist}")
        manual_window.geometry("600x400")

        # Create a TreeView to display the top 10 fuzzy matches
        tree = ttk.Treeview(manual_window, columns=("Track", "Artist", "Album", "Score", "RatingKey"), show="headings")
        tree.heading("Track", text="Track")
        tree.heading("Artist", text="Artist")
        tree.heading("Album", text="Album")
        tree.heading("Score", text="Score")
        tree.heading("RatingKey", text="RatingKey")
        tree.pack(fill="both", expand=True)

        # Query the database for potential matches based on track name and artist
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Get all tracks from the database to compare
        cursor.execute("""
            SELECT 
                mi.title AS track_name,
                CASE 
                    WHEN mi.original_title IS NOT NULL AND mi.original_title != '' THEN mi.original_title
                    ELSE (SELECT title FROM metadata_items WHERE id = mi.parent_id)
                END AS artist_name,
                (SELECT title FROM metadata_items WHERE id = mi.parent_id) AS album_name,
                mi.id AS rating_key
            FROM metadata_items mi
            WHERE mi.metadata_type = 10
        """)
        results = cursor.fetchall()
    
        # Perform fuzzy matching and score calculation
        matches = []
        for result in results:
            result_track, result_artist, result_album, rating_key = result
            track_score = fuzz.token_set_ratio(track, result_track)
            artist_score = fuzz.token_set_ratio(artist, result_artist)
            total_score = (track_score + artist_score) // 2  # Average score between track and artist
    
            matches.append((result_track, result_artist, result_album, total_score, rating_key))
    
        # Sort matches by score in descending order and take top 10 matches
        matches = sorted(matches, key=lambda x: x[3], reverse=True)[:10]

        # Insert the matches into the TreeView
        for match in matches:
            tree.insert("", "end", values=(match[0], match[1], match[2], match[3], match[4]))
    
        conn.close()

        # Create a button to confirm the selection
        def confirm_selection():
            selected_item = tree.selection()
            if selected_item:
                selected_track_info = tree.item(selected_item[0], "values")
                # Update the main TreeView with the selected match
                self.results_frame.item(item_id, values=(
                    track, artist, selected_track_info[0], selected_track_info[1], selected_track_info[2], selected_track_info[3], selected_track_info[4]))
                # Update the track map with the correct match
                self.track_map[item_id] = (selected_track_info[0], selected_track_info[1], selected_track_info[2], selected_track_info[3], int(selected_track_info[4]))
                manual_window.destroy()
    
        confirm_button = tk.Button(manual_window, text="Confirm", command=confirm_selection)
        confirm_button.pack(pady=10)

    def load_database(self):
        """Select and load the Plex database file."""
        self.db_path = filedialog.askopenfilename(title="Select Database File", filetypes=[("SQLite Files", "*.db")])
        if self.db_path:
            self.go_button.config(state=tk.NORMAL)
            messagebox.showinfo("Success", f"Database file '{self.db_path}' loaded successfully!")
 

    def load_csv(self):
        """Load CSV file and parse it."""
        file_path = filedialog.askopenfilename(title="Select Tracklist CSV", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.csv_data = pd.read_csv(file_path, sep="\t", header=None, names=["Track", "Artist"])
            self.results_frame.delete(*self.results_frame.get_children())
            self.progress_label.config(text=f"Progress: 0/{len(self.csv_data)}")

    def process_tracks(self):
        """Search and match each track in the CSV file against the Plex database."""
        if not self.db_path:
            messagebox.showwarning("Error", "No database file selected!")
            return

        self.results = []
        total_tracks = len(self.csv_data)

        # Connect to the SQLite database
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        for index, row in self.csv_data.iterrows():
            track = row["Track"]
            artist = row["Artist"]

            # Run the recursive SQL query to search by track name and return the metadata_items.id
            cursor.execute(self.get_recursive_query(), (f"%{track}%",))
            track_results = cursor.fetchall()

            best_match = self.fuzzy_match(track_results, track, artist)

            if best_match and best_match[4]:  # Check if best_match and it has a valid RatingKey
                match_score = best_match[3] if len(best_match) > 3 else 0
                tag = 'amber' if match_score < 100 else ''
                track_id = self.results_frame.insert("", "end", values=(
                    track,
                    artist,
                    best_match[0] if len(best_match) > 0 else "No Match",
                    best_match[1] if len(best_match) > 1 else "No Match",
                    best_match[2] if len(best_match) > 2 else "No Match",
                    best_match[3] if len(best_match) > 3 else "0",
                    best_match[4] if len(best_match) > 4 else "N/A"
                ), tags=(tag,))
                self.track_map[track_id] = best_match  # Store the whole tuple
            else:
                # No match, mark as amber
                track_id = self.results_frame.insert("", "end", values=(track, artist, "No Match", "", "", "0", "N/A"), tags=('amber',))
                self.track_map[track_id] = {"track": track, "artist": artist}  # Store the track and artist for later manual matching

            # Update progress
            self.progress_label.config(text=f"Progress: {index + 1}/{total_tracks}")
            self.root.update_idletasks()

        conn.close()
        self.save_button.config(state=tk.NORMAL)

    def get_recursive_query(self):
        """Recursive SQL query to search for tracks by name and return metadata_items.id."""
        return """
        WITH RECURSIVE hierarchy AS (
            SELECT 
                id,
                parent_id,
                title,
                original_title,
                metadata_type
            FROM 
                metadata_items
            WHERE 
                title LIKE ? AND metadata_type = 10
            UNION ALL
            SELECT 
                mi.id,
                mi.parent_id,
                mi.title,
                mi.original_title,
                mi.metadata_type
            FROM 
                metadata_items mi
            INNER JOIN 
                hierarchy h ON mi.id = h.parent_id
        )
        SELECT 
            h.title AS track_name,
            CASE 
                WHEN h.original_title IS NOT NULL AND h.original_title != '' THEN h.original_title
                ELSE (SELECT title FROM metadata_items WHERE id = (SELECT parent_id FROM metadata_items WHERE id = h.id))
            END AS true_track_artist,
            (SELECT title FROM metadata_items WHERE id = h.parent_id) AS album,
            h.id  -- Ensure that we return the metadata_items.id
        FROM 
            hierarchy h
        WHERE 
            h.metadata_type = 10;
        """

    def fuzzy_match(self, results, track, artist):
        """Perform fuzzy matching on the LIKE results."""
        best_match = None
        highest_score = 0

        # Loop over the results to find the best match based on fuzzy score
        for result in results:
            if len(result) < 4:
                continue  # Skip if the result doesn't have enough elements
            
            result_track, result_artist, result_album, result_id = result
            score = fuzz.token_set_ratio(f"{track} {artist}", f"{result_track} {result_artist}")
            
            if score > highest_score:
                highest_score = score
                best_match = (result_track, result_artist, result_album, score, result_id)
        
        # If no best match is found, return a default tuple
        if not best_match:
            return ("No Match", "No Match", "No Match", 0, None)
        
        return best_match

    def save_playlist(self):
        """Create a playlist from the selected tracks using the reliable method."""
        playlist_name = simpledialog.askstring("Playlist Name", "Enter the name for the playlist:")
        if not playlist_name:
            return

        # Get the selected items from the TreeView
        selected_items = self.results_frame.selection()  # Get selected rows in the TreeView

        selected_tracks = []
        for item_id in selected_items:
            match = self.track_map.get(item_id)
            if match:
                rating_key = match[4]  # metadata_items.id (RatingKey) is stored at index 4

                # Ensure the rating_key is properly formatted and a valid integer
                if rating_key and isinstance(rating_key, int):
                    try:
                        # Fetch the track directly using the RatingKey
                        plex_track = plex.fetchItem(rating_key)
                        if plex_track:
                            selected_tracks.append(plex_track)
                        else:
                            print(f"Track not found in Plex with RatingKey: {rating_key}")
                    except requests.exceptions.InvalidURL as e:
                        print(f"Invalid URL: {e}")
                        messagebox.showerror("Error", f"Invalid URL encountered for RatingKey: {rating_key}")
                else:
                    print(f"Invalid RatingKey: {rating_key}")

        # If we have selected tracks, create the playlist
        if selected_tracks:
            try:
                # Use the reliable playlist creation method from the provided script
                Playlist.create(server=plex, title=playlist_name, items=selected_tracks)
                messagebox.showinfo("Success", f"Playlist '{playlist_name}' created successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create playlist: {str(e)}")
        else:
            messagebox.showwarning("Error", "No tracks selected for the playlist.")

    def fix_selected_row(self):
        """Allow the user to manually select a match from second query results."""
        selected_item = self.results_frame.selection()[0]
        track_info = self.track_map.get(selected_item)

        if track_info and isinstance(track_info, dict):  # We only want to fix rows with no match
            track = track_info["track"]
            artist = track_info["artist"]

            # Perform second query to search by artist and show results for manual selection
            self.show_manual_selection_window(track, artist, selected_item)

    def show_manual_selection_window(self, track, artist, item_id):
        """Show a popup window to allow manual selection of a match."""
        # Create a popup window
        manual_window = Toplevel(self.root)
        manual_window.title(f"Manual Match for {track} by {artist}")
        manual_window.geometry("600x400")

        # Create a TreeView to display results
        tree = ttk.Treeview(manual_window, columns=("Track", "Artist", "Album", "RatingKey"), show="headings")
        tree.heading("Track", text="Track")
        tree.heading("Artist", text="Artist")
        tree.heading("Album", text="Album")
        tree.heading("RatingKey", text="RatingKey")
        tree.pack(fill="both", expand=True)

        # Query the database for potential matches by artist
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        cursor.execute(self.get_artist_query(), (f"%{artist}%",))
        results = cursor.fetchall()

        for result in results:
            tree.insert("", "end", values=(result[0], result[1], result[2], result[3]))

        conn.close()

        # Create a button to confirm the selection
        def confirm_selection():
            selected_item = tree.selection()
            if selected_item:
                selected_track_info = tree.item(selected_item[0], "values")
                # Update the main TreeView with the selected match
                self.results_frame.item(item_id, values=(track, artist, selected_track_info[0], selected_track_info[1], selected_track_info[2], "100", selected_track_info[3]))
                # Update the track map with the correct match
                self.track_map[item_id] = (selected_track_info[0], selected_track_info[1], selected_track_info[2], 100, int(selected_track_info[3]))
                manual_window.destroy()

        confirm_button = tk.Button(manual_window, text="Confirm", command=confirm_selection)
        confirm_button.pack(pady=10)

    def get_artist_query(self):
        """Query to search for tracks by artist."""
        return """
        SELECT 
            h.title AS track_name,
            CASE 
                WHEN h.original_title IS NOT NULL AND h.original_title != '' THEN h.original_title
                ELSE (SELECT title FROM metadata_items WHERE id = (SELECT parent_id FROM metadata_items WHERE id = h.id))
            END AS true_track_artist,
            (SELECT title FROM metadata_items WHERE id = h.parent_id) AS album,
            h.id AS rating_key
        FROM 
            metadata_items h
        WHERE 
            h.metadata_type = 10 AND h.original_title LIKE ?
        """

    def on_row_click(self, event):
        """Handle click event to show popup if a manual selection is required."""
        pass

    def show_context_menu(self, event):
        """Show context menu for right-click to remove, fix, or fuzzy match a row."""
        selected_item = self.results_frame.identify_row(event.y)
        if selected_item:
            self.results_frame.selection_set(selected_item)
            self.context_menu.post(event.x_root, event.y_root)

    def remove_selected_row(self):
        """Remove the selected row from the TreeView."""
        selected_item = self.results_frame.selection()[0]
        self.results_frame.delete(selected_item)
        del self.track_map[selected_item]


if __name__ == "__main__":
    root = tk.Tk()
    app = PlaylistApp(root)
    root.mainloop()
