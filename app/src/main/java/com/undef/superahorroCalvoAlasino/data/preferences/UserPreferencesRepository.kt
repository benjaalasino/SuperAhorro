package com.undef.superahorroCalvoAlasino.data.preferences

import android.content.Context
import androidx.datastore.preferences.core.booleanPreferencesKey
import androidx.datastore.preferences.core.edit
import androidx.datastore.preferences.core.stringPreferencesKey
import kotlinx.coroutines.flow.Flow
import kotlinx.coroutines.flow.first
import kotlinx.coroutines.flow.map

class UserPreferencesRepository(context: Context) {

    private val dataStore = context.userPrefs

    companion object {
        private val IS_LOGGED_IN = booleanPreferencesKey("IS_LOGGED_IN")
        private val USER_EMAIL = stringPreferencesKey("USER_EMAIL")
        private val USER_NAME = stringPreferencesKey("USER_NAME")
        private val DARK_MODE = booleanPreferencesKey("DARK_MODE")
        private val LAST_SORT_ORDER = stringPreferencesKey("LAST_SORT_ORDER")
    }

    val isLoggedIn: Flow<Boolean> = dataStore.data.map { prefs ->
        prefs[IS_LOGGED_IN] ?: false
    }

    val userEmail: Flow<String> = dataStore.data.map { prefs ->
        prefs[USER_EMAIL] ?: ""
    }

    val userName: Flow<String> = dataStore.data.map { prefs ->
        prefs[USER_NAME] ?: ""
    }

    val darkMode: Flow<Boolean> = dataStore.data.map { prefs ->
        prefs[DARK_MODE] ?: false
    }

    val lastSortOrder: Flow<String> = dataStore.data.map { prefs ->
        prefs[LAST_SORT_ORDER] ?: "fecha_desc"
    }

    suspend fun saveSession(email: String, name: String) {
        dataStore.edit { prefs ->
            prefs[IS_LOGGED_IN] = true
            prefs[USER_EMAIL] = email
            prefs[USER_NAME] = name
        }
    }

    suspend fun clearSession() {
        dataStore.edit { prefs ->
            prefs[IS_LOGGED_IN] = false
            prefs.remove(USER_EMAIL)
            prefs.remove(USER_NAME)
        }
    }

    suspend fun setDarkMode(enabled: Boolean) {
        dataStore.edit { prefs ->
            prefs[DARK_MODE] = enabled
        }
    }

    suspend fun setLastSortOrder(order: String) {
        dataStore.edit { prefs ->
            prefs[LAST_SORT_ORDER] = order
        }
    }

    // Lee el valor real desde DataStore (sin el initial value del StateFlow)
    suspend fun readInitialLoginState(): Boolean = isLoggedIn.first()
}
