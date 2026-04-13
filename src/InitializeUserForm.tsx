import { useState } from 'react';
import type { User } from './types';
import {bgiurl} from "./url.ts";

interface InitializeUserFormProps {
    users: User[];
}

export function InitializeUserForm({ users }: InitializeUserFormProps) {
    const [showUserList, setShowUserList] = useState(false);
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);
    const [success, setSuccess] = useState(false);

    const handleInitializeUser = async (email: string) => {
        setLoading(true);
        setError(null);
        setSuccess(false);

        try {
            const response = await fetch(`${bgiurl}/users/initialize`, {
                method: 'PUT',
                headers: {
                    'Content-Type': 'application/json',
                },
                body: JSON.stringify({ emailAddress: email }),
            });

            if (!response.ok) {
                const contentType = response.headers.get('content-type');
                let errorMessage = response.statusText;

                if (contentType?.includes('application/json')) {
                    const errorData = await response.json();
                    errorMessage = errorData.message || errorMessage;
                } else {
                    errorMessage = await response.text();
                }

                console.log(errorMessage);
                throw new Error(errorMessage || 'Failed to initialize user');
            }

            setSuccess(true);
            setShowUserList(false);

            // Hide success message after 3 seconds
            setTimeout(() => {
                setSuccess(false);
            }, 3000);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'An error occurred');
        } finally {
            setLoading(false);
        }
    };

    return (
        <div>
            {/*<button*/}
            {/*    onClick={() => setShowUserList(!showUserList)}*/}
            {/*    disabled={loading}*/}
            {/*    style={{*/}
            {/*        width: '100%',*/}
            {/*        padding: '10px 12px',*/}
            {/*        border: 'none',*/}
            {/*        background: 'transparent',*/}
            {/*        cursor: loading ? 'not-allowed' : 'pointer',*/}
            {/*        textAlign: 'left',*/}
            {/*        borderRadius: 6,*/}
            {/*    }}*/}
            {/*    onMouseEnter={(e) => !loading && (e.currentTarget.style.background = '#f5f5f5')}*/}
            {/*    onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}*/}
            {/*>*/}
            {/*    Initialize User*/}
            {/*</button>*/}

            {showUserList && users.length > 0 && (
                <div style={{
                    maxHeight: 200,
                    overflow: 'auto',
                    borderTop: '1px solid #eee',
                    marginTop: 4,
                    paddingTop: 4
                }}>
                    {users.map(user => (
                        <button
                            key={user.email}
                            onClick={() => handleInitializeUser(user.email)}
                            disabled={loading}
                            style={{
                                width: '100%',
                                padding: '8px 12px',
                                border: 'none',
                                background: 'transparent',
                                cursor: loading ? 'not-allowed' : 'pointer',
                                textAlign: 'left',
                                borderRadius: 6,
                                fontSize: 14
                            }}
                            onMouseEnter={(e) => !loading && (e.currentTarget.style.background = '#f5f5f5')}
                            onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
                        >
                            {user.email}
                        </button>
                    ))}
                </div>
            )}

            {error && <div style={{color: 'red', padding: '8px 12px', fontSize: 12}}>{error}</div>}
            {success && <div style={{color: 'green', padding: '8px 12px', fontSize: 12}}>✓ User initialized!</div>}
        </div>
    );
}