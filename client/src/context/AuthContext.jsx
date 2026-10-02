import React, { createContext, useContext, useState, useEffect } from 'react';
import { useMsal } from "@azure/msal-react";
import { loginRequest } from "../config/authConfig";
import axios from 'axios';

const AuthContext = createContext();

export const useAuth = () => useContext(AuthContext);

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:5000/api';
const BACKEND_URL = `${API_BASE_URL}/auth`;

export const AuthProvider = ({ children }) => {
    const { instance } = useMsal();
    const [currentUser, setCurrentUser] = useState(null);
    const [loading, setLoading] = useState(true);

    const checkAuthStatus = async () => {
        try {
            const { data } = await axios.get(`${BACKEND_URL}/is-auth`, { withCredentials: true });
            if (data.success) {
                setCurrentUser(data.userData);
            } else {
                setCurrentUser(null);
            }
        } catch (error) {
            setCurrentUser(null);
        } finally {
            setLoading(false);
        }
    };

    useEffect(() => {
        checkAuthStatus();
    }, []);

    const login = async () => {
        try {
            // 1. Open Microsoft Login Popup
            const loginResponse = await instance.loginPopup(loginRequest);
            const account = loginResponse.account;
            
            // 2. Extract email safely
            const userEmail = account?.username || account?.idTokenClaims?.preferred_username || account?.idTokenClaims?.email || "";

            // 3. FRONTEND VALIDATION (Case-Insensitive)
            if (!userEmail.toLowerCase().endsWith('@gst.sies.edu.in')) {
                // If it fails, force clear this specific account from cache immediately
                await instance.logoutPopup({ account: account });
                // Throw specific error to be caught by Login.jsx
                throw new Error("INVALID_DOMAIN");
            }

            // 4. If domain is valid, proceed to backend authentication
            const accessToken = loginResponse.accessToken;
            const { data } = await axios.post(
                `${BACKEND_URL}/microsoft-login`,
                { accessToken },
                { withCredentials: true }
            );

            if (data.success) {
                setCurrentUser(data.userData);
                return data.userData;
            } else {
                throw new Error(data.message);
            }
        } catch (error) {
            throw error;
        }
    };

    const logout = async () => {
        try {
            await axios.post(`${BACKEND_URL}/logout`, {}, { withCredentials: true });
            await instance.logoutPopup();
            setCurrentUser(null);
        } catch (error) {
            console.error("Logout Error:", error);
        }
    };

    return (
        <AuthContext.Provider value={{ 
            currentUser, 
            setCurrentUser,
            login, 
            logout, 
            loading
        }}>
            {!loading && children}
        </AuthContext.Provider>
    );
};