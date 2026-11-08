import { useState } from 'react';
import './InstitutionalHeader.css';

export default function InstitutionalHeader() {
  const [nursingError, setNursingError] = useState(false);
  const [azharError, setAzharError] = useState(false);
  const [deptError, setDeptError] = useState(false);

  return (
    <header className="inst-header" role="banner" aria-label="Institutional Identity Header">
      {/* Left: Faculty of Nursing for Girls, Al-Azhar University */}
      <div className="inst-side inst-side-left">
        <div className="inst-logo-frame inst-logo-frame-nursing" title="Faculty of Nursing for Girls - Al-Azhar University">
          {!nursingError ? (
            <img
              src={`${import.meta.env.BASE_URL}images/nursing-college-logo.png`}
              alt="Faculty of Nursing for Girls, Al-Azhar University Official Logo"
              className="inst-logo"
              onError={() => setNursingError(true)}
            />
          ) : (
            <div className="inst-logo-fallback">
              <span>AFN</span>
            </div>
          )}
        </div>
        <div className="inst-text-block">
          <span className="inst-en-name">Faculty of Nursing for Girls, Al-Azhar University</span>
          <span className="inst-ar-name">كُليّة التمريض للبنات بالقاهرة</span>
        </div>
      </div>

      {/* Center: Department Identity */}
      <div className="inst-center">
        {!deptError ? (
          <div className="inst-dept-block" title="Medical-Surgical Nursing Department">
            <div className="inst-logo-frame inst-logo-frame-dept">
              <img
                src={`${import.meta.env.BASE_URL}images/dept-image.png`}
                alt="Medical-Surgical Nursing Department Logo"
                className="inst-logo"
                onError={() => setDeptError(true)}
              />
            </div>
            <div className="inst-text-block">
              <span className="inst-en-name">Medical-Surgical Nursing Department</span>
              <span className="inst-ar-name">قسم التمريض الباطني والجراحي</span>
            </div>
          </div>
        ) : (
          <div className="inst-topic-badge">
            <span className="inst-dept-tag">قسم التمريض الباطني والجراحي</span>
            <span className="inst-dot" aria-hidden="true">•</span>
            <span className="inst-topic-name">COLON CANCER</span>
          </div>
        )}
      </div>

      {/* Right: Al-Azhar University */}
      <div className="inst-side inst-side-right">
        <div className="inst-text-block inst-text-right">
          <span className="inst-en-name">Al-Azhar University</span>
          <span className="inst-ar-name">جامعة الأزهر الشريف</span>
        </div>
        <div className="inst-logo-frame inst-logo-frame-azhar" title="Al-Azhar University">
          {!azharError ? (
            <img
              src={`${import.meta.env.BASE_URL}images/al-azhar-university-logo.png`}
              alt="Al-Azhar University Official Logo"
              className="inst-logo"
              onError={() => setAzharError(true)}
            />
          ) : (
            <div className="inst-logo-fallback">
              <span>AZHAR</span>
            </div>
          )}
        </div>
      </div>
    </header>
  );
}
