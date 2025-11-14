"""REST API server for FileManagerApp"""
import os
import threading
from flask import Flask, jsonify, request


class FileManagerAPI:
    """REST API server for FileManagerApp"""
    
    def __init__(self, app_instance, host='127.0.0.1', port=5000):
        self.app = app_instance
        self.host = host
        self.port = port
        self.flask_app = Flask(__name__)
        self.server_thread = None
        self.setup_routes()
    
    def setup_routes(self):
        """Setup all API routes"""
        
        @self.flask_app.route('/api/tracking/start', methods=['POST'])
        def start_tracking():
            try:
                # Use root.after to ensure thread-safe GUI operations
                self.app.root.after(0, self.app.start_tracking)
                return jsonify({'success': True, 'message': 'Tracking started'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/tracking/stop', methods=['POST'])
        def stop_tracking():
            try:
                self.app.root.after(0, self.app.stop_tracking)
                return jsonify({'success': True, 'message': 'Tracking stopped'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/copy_rename', methods=['POST'])
        def copy_rename():
            try:
                self.app.root.after(0, self.app.copy_and_rename)
                return jsonify({'success': True, 'message': 'Copy and rename initiated'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/source_path', methods=['GET', 'POST'])
        def source_path():
            try:
                if request.method == 'GET':
                    path = self.app.source_folder.get()
                    return jsonify({'success': True, 'path': path})
                else:  # POST
                    data = request.get_json()
                    if 'path' not in data:
                        return jsonify({'success': False, 'error': 'Path parameter required'}), 400
                    self.app.root.after(0, lambda: self.app.source_folder.set(data['path']))
                    self.app.root.after(0, self.app.update_path_labels)
                    self.app.root.after(0, self.app.update_button_states)
                    return jsonify({'success': True, 'message': 'Source path updated'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/destination_path', methods=['GET', 'POST'])
        def destination_path():
            try:
                if request.method == 'GET':
                    path = self.app.dest_folder.get()
                    return jsonify({'success': True, 'path': path})
                else:  # POST
                    data = request.get_json()
                    if 'path' not in data:
                        return jsonify({'success': False, 'error': 'Path parameter required'}), 400
                    self.app.root.after(0, lambda: self.app.dest_folder.set(data['path']))
                    self.app.root.after(0, self.app.update_path_labels)
                    self.app.root.after(0, self.app.update_button_states)
                    return jsonify({'success': True, 'message': 'Destination path updated'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/name_pattern', methods=['GET', 'POST'])
        def name_pattern():
            try:
                if request.method == 'GET':
                    pattern = self.app.naming_pattern.get()
                    return jsonify({'success': True, 'pattern': pattern})
                else:  # POST
                    data = request.get_json()
                    if 'pattern' not in data:
                        return jsonify({'success': False, 'error': 'Pattern parameter required'}), 400
                    self.app.root.after(0, lambda: self.app.naming_pattern.set(data['pattern']))
                    return jsonify({'success': True, 'message': 'Name pattern updated'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/tracking', methods=['GET', 'DELETE'])
        def tracking():
            try:
                if request.method == 'GET':
                    count = request.args.get('count', default=10, type=int)
                    if count < 1:
                        count = 1
                    if count > 1000:  # Reasonable limit
                        count = 1000
                    
                    # Get the latest files (last N files in the list)
                    tracked_files = self.app.tracked_files
                    latest_files = tracked_files[-count:] if len(tracked_files) > count else tracked_files
                    
                    result = []
                    for i, file_path in enumerate(latest_files):
                        # Find the original index in the full list
                        original_index = self.app.tracked_files.index(file_path)
                        original_name = os.path.basename(file_path)
                        preview_name = self.app.generate_filename_preview(original_index, len(self.app.tracked_files))
                        file_ext = os.path.splitext(file_path)[1]
                        new_name = preview_name + file_ext
                        
                        # Check file state (color coding)
                        has_duplicate = self.app.has_duplicate_preview_name(original_index, new_name)
                        exists_in_dest = self.app.preview_exists_in_destination(new_name)
                        
                        # Determine state
                        if has_duplicate:
                            state = 'duplicate'  # Red - collides with another tracked file
                        elif exists_in_dest:
                            state = 'exists'  # Blue - already exists in destination
                        else:
                            state = 'normal'  # No conflicts
                        
                        result.append({
                            'original_path': file_path,
                            'original_name': original_name,
                            'new_name': new_name,
                            'state': state
                        })
                    
                    return jsonify({
                        'success': True,
                        'files': result,
                        'total_count': len(tracked_files),
                        'returned_count': len(result)
                    })
                else:  # DELETE
                    self.app.root.after(0, self.app.clear_tracked)
                    return jsonify({'success': True, 'message': 'Tracked files cleared'})
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
        
        @self.flask_app.route('/api/status', methods=['GET'])
        def status():
            """Get current status of the application"""
            try:
                is_tracking = self.app.observer is not None
                tracked_count = len(self.app.tracked_files)
                return jsonify({
                    'success': True,
                    'is_tracking': is_tracking,
                    'tracked_files_count': tracked_count,
                    'source_path': self.app.source_folder.get(),
                    'destination_path': self.app.dest_folder.get(),
                    'name_pattern': self.app.naming_pattern.get()
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)}), 500
    
    def start(self):
        """Start the API server in a background thread"""
        if self.server_thread is None or not self.server_thread.is_alive():
            self.server_thread = threading.Thread(target=self._run_server, daemon=True)
            self.server_thread.start()
            print(f"API server started on http://{self.host}:{self.port}")
    
    def _run_server(self):
        """Run the Flask server (called in background thread)"""
        self.flask_app.run(host=self.host, port=self.port, debug=False, use_reloader=False)
    
    def stop(self):
        """Stop the API server"""
        # Flask doesn't have a clean way to stop, but since it's a daemon thread,
        # it will be terminated when the main process exits
        pass

