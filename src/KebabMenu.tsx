import { useEffect, useRef, useState } from "react";
import {ConnectGoogleButton} from "./ConnectGoogleButton.tsx";
// import {InitializeUserForm} from "./InitializeUserForm.tsx";
import type { User } from "./types.ts";

interface KebabMenuProps {
    users: User[];
    selectedUser: string;
    onSelectUser: (email: string) => void;
    onCreateUser: () => void;
    onCreateTeam: () => void;
    onStartSession: () => void;
    onEndSession: () => void;
}

export function KebabMenu({ users, selectedUser, onSelectUser, onCreateUser, onCreateTeam, onStartSession, onEndSession }: KebabMenuProps) {
    const [open, setOpen] = useState(false);
    const [showUserList, setShowUserList] = useState(false);
    const [showEndSessionConfirm, setShowEndSessionConfirm] = useState(false);
    const [endSessionInput, setEndSessionInput] = useState('');
    const wrapRef = useRef<HTMLDivElement | null>(null);

    useEffect(() => {
        function onDown(e: MouseEvent) {
            if (!wrapRef.current) return;
            if (!wrapRef.current.contains(e.target as Node)) {
                setOpen(false);
                setShowUserList(false);
                setShowEndSessionConfirm(false);
                setEndSessionInput('');
            }
        }
        function onKey(e: KeyboardEvent) {
            if (e.key === "Escape") {
                setOpen(false);
                setShowUserList(false);
                setShowEndSessionConfirm(false);
                setEndSessionInput('');
            }
        }
        document.addEventListener("mousedown", onDown);
        document.addEventListener("keydown", onKey);
        return () => {
            document.removeEventListener("mousedown", onDown);
            document.removeEventListener("keydown", onKey);
        };
    }, []);

    const handleSelectUser = (email: string) => {
        onSelectUser(email);
        setShowUserList(false);
        setOpen(false);
    };

    return (
        <div ref={wrapRef} style={{ position: "relative", display: "inline-block" }}>
            <button
                type="button"
                aria-haspopup="menu"
                aria-expanded={open}
                onClick={() => setOpen((v) => !v)}
                style={{
                    width: 36,
                    height: 36,
                    borderRadius: 999,
                    border: "1px solid #00000022",
                    background: "#fff",
                    cursor: "pointer",
                    display: "grid",
                    placeItems: "center",
                    padding: 0,
                }}
            >
                <span style={{ fontSize: 18, lineHeight: 1, letterSpacing: 2 }}>⋯</span>
            </button>

            {open && (
                <div
                    role="menu"
                    style={{
                        position: "absolute",
                        right: 0,
                        top: "calc(100% + 8px)",
                        minWidth: 200,
                        background: "#fff",
                        border: "1px solid #00000022",
                        borderRadius: 12,
                        boxShadow: "0 10px 30px rgba(0,0,0,0.12)",
                        padding: 6,
                        zIndex: 1000,
                        overflow: "hidden"
                    }}
                >
                    <div style={{ borderBottom: '1px solid #ddd', marginBottom: 6, paddingBottom: 6 }}>
                        <ConnectGoogleButton/>
                    </div>

                    {/*<div style={{ borderBottom: '1px solid #ddd', marginBottom: 6, paddingBottom: 6 }}>*/}
                    {/*    <InitializeUserForm users={users} />*/}
                    {/*</div>*/}

                    <button
                        onClick={() => setShowUserList(!showUserList)}
                        style={{
                            width: '100%',
                            padding: '10px 12px',
                            border: '1px solid #ddd',
                            background: 'transparent',
                            cursor: 'pointer',
                            textAlign: 'left',
                            borderRadius: 6,
                            marginBottom: 6,
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.background = '#f5f5f5'}
                        onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                    >
                        Choose Brain {selectedUser && `(${selectedUser})`}
                    </button>

                    {showUserList && users.length > 0 && (
                        <div style={{
                            maxHeight: 200,
                            overflow: 'auto',
                            border: '1px solid #ddd',
                            borderRadius: 6,
                            marginBottom: 6,
                            padding: 4
                        }}>
                            {[...users].sort((a, b) => a.email.localeCompare(b.email)).map(user => (
                                <button
                                    key={user.email}
                                    onClick={() => handleSelectUser(user.email)}
                                    style={{
                                        width: '100%',
                                        padding: '8px 12px',
                                        border: '1px solid #e0e0e0',
                                        background: selectedUser === user.email ? '#e3f2fd' : 'transparent',
                                        cursor: 'pointer',
                                        textAlign: 'left',
                                        borderRadius: 6,
                                        fontSize: 14,
                                        marginBottom: 4
                                    }}
                                    onMouseEnter={(e) => e.currentTarget.style.background = selectedUser === user.email ? '#e3f2fd' : '#f5f5f5'}
                                    onMouseLeave={(e) => e.currentTarget.style.background = selectedUser === user.email ? '#e3f2fd' : 'transparent'}
                                >
                                    {user.email}
                                </button>
                            ))}
                        </div>
                    )}

                    <button
                        onClick={() => {
                            onCreateUser();
                            setOpen(false);
                        }}
                        style={{
                            width: '100%',
                            padding: '10px 12px',
                            border: '1px solid #ddd',
                            background: 'transparent',
                            cursor: 'pointer',
                            textAlign: 'left',
                            borderRadius: 6,
                            marginBottom: 6,
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.background = '#f5f5f5'}
                        onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                    >
                        + Create New Brain
                    </button>

                    <button
                        onClick={() => {
                            onCreateTeam();
                            setOpen(false);
                        }}
                        style={{
                            width: '100%',
                            padding: '10px 12px',
                            border: '1px solid #ddd',
                            background: 'transparent',
                            cursor: 'pointer',
                            textAlign: 'left',
                            borderRadius: 6,
                            marginBottom: 6,
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.background = '#f5f5f5'}
                        onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                    >
                        + Create New Team
                    </button>

                    <button
                        onClick={() => {
                            onStartSession();
                            setOpen(false);
                        }}
                        style={{
                            width: '100%',
                            padding: '10px 12px',
                            border: '1px solid #ddd',
                            background: 'transparent',
                            cursor: 'pointer',
                            textAlign: 'left',
                            borderRadius: 6,
                            marginBottom: 6,
                            color: '#28a745',
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.background = '#f5f5f5'}
                        onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                    >
                        ▶ Start Session
                    </button>

                    <button
                        onClick={() => setShowEndSessionConfirm(true)}
                        style={{
                            width: '100%',
                            padding: '10px 12px',
                            border: '1px solid #ddd',
                            background: 'transparent',
                            cursor: 'pointer',
                            textAlign: 'left',
                            borderRadius: 6,
                            color: '#dc3545',
                        }}
                        onMouseEnter={(e) => e.currentTarget.style.background = '#f5f5f5'}
                        onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                    >
                        ⏹ End Session
                    </button>

                    {showEndSessionConfirm && (
                        <div style={{
                            border: '1px solid #dc3545',
                            borderRadius: 6,
                            padding: 10,
                            marginTop: 6,
                            background: '#fff5f5',
                        }}>
                            <p style={{ margin: '0 0 8px 0', fontSize: 13, color: '#dc3545', fontWeight: 600 }}>
                                Type <em>End Session</em> to confirm:
                            </p>
                            <input
                                type="text"
                                value={endSessionInput}
                                onChange={(e) => setEndSessionInput(e.target.value)}
                                placeholder="End Session"
                                autoFocus
                                style={{
                                    width: '100%',
                                    padding: '6px 8px',
                                    fontSize: 13,
                                    border: '1px solid #ccc',
                                    borderRadius: 4,
                                    marginBottom: 8,
                                    boxSizing: 'border-box',
                                }}
                            />
                            <div style={{ display: 'flex', gap: 6 }}>
                                <button
                                    onClick={() => {
                                        if (endSessionInput === 'End Session') {
                                            onEndSession();
                                            setOpen(false);
                                            setShowEndSessionConfirm(false);
                                            setEndSessionInput('');
                                        }
                                    }}
                                    disabled={endSessionInput !== 'End Session'}
                                    style={{
                                        flex: 1,
                                        padding: '6px 10px',
                                        fontSize: 13,
                                        background: endSessionInput === 'End Session' ? '#dc3545' : '#ccc',
                                        color: 'white',
                                        border: 'none',
                                        borderRadius: 4,
                                        cursor: endSessionInput === 'End Session' ? 'pointer' : 'not-allowed',
                                    }}
                                >
                                    Confirm
                                </button>
                                <button
                                    onClick={() => {
                                        setShowEndSessionConfirm(false);
                                        setEndSessionInput('');
                                    }}
                                    style={{
                                        flex: 1,
                                        padding: '6px 10px',
                                        fontSize: 13,
                                        background: '#6c757d',
                                        color: 'white',
                                        border: 'none',
                                        borderRadius: 4,
                                        cursor: 'pointer',
                                    }}
                                >
                                    Cancel
                                </button>
                            </div>
                        </div>
                    )}
                </div>
            )}
        </div>
    );
}