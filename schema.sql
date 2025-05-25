CREATE TABLE users (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        username TEXT UNIQUE NOT NULL,
                        password TEXT NOT NULL,
                        name TEXT NOT NULL,
                        email TEXT NOT NULL,
                        role TEXT NOT NULL DEFAULT 'staff'
                    );
CREATE TABLE sqlite_sequence(name,seq);
CREATE TABLE requirements (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        title TEXT NOT NULL,
                        description TEXT NOT NULL,
                        assigner_id INTEGER NOT NULL,
                        assignee_id INTEGER NOT NULL,
                        status TEXT NOT NULL DEFAULT 'pending',
                        priority TEXT NOT NULL DEFAULT 'normal',
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        scheduled_time TIMESTAMP,
                        is_dispatched INTEGER DEFAULT 1,
                        completed_at TIMESTAMP,
                        comment TEXT,
                        is_deleted INTEGER DEFAULT 0,
                        deleted_at TIMESTAMP,
                        FOREIGN KEY (assigner_id) REFERENCES users (id),
                        FOREIGN KEY (assignee_id) REFERENCES users (id)
                    );
