/*
 * NoRunningException.java
 *
 * Created on 2012-01-07
 * 
 * Copyright (c) 2012 imacat
 */

/*
 * Copyright (c) 2012 imacat
 * 
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * 
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */
package tw.idv.imacat.android.mpresent.uno;

/**
 * Is thrown when there is no running presentation.
 *
 * @author <a href="mailto:imacat&#64;mail.imacat.idv.tw">imacat</a>
 * @version 0.1.0
 */
public class NoRunningException extends Exception {
    
    /**
     * Constructs an instance of @NoRunningException with the specified
     * detail message.
     *
     * @param message the detail message
     */
    public NoRunningException(String message) {
        super(message);
        return;
    }
}
