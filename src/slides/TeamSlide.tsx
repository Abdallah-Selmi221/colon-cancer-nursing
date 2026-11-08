import { Users, Award, GraduationCap } from 'lucide-react';
import type { Slide } from '../content/slides';

const STUDENTS = [
  { name: 'Hager Hany', role: 'Nursing Student' },
  { name: 'Hager Saeed', role: 'Nursing Student' },
  { name: 'Hager Mostafa', role: 'Nursing Student' },
  { name: 'Haidy Mahmoud', role: 'Nursing Student' },
  { name: 'Norhan Badawy', role: 'Nursing Student' },
  { name: 'Hager Ahmed', role: 'Nursing Student' },
];

export default function TeamSlide({ slide: _slide }: { slide: Slide }) {
  return (
    <div className="slide-body team-slide-body">
      <div className="team-header-row" data-reveal="head">
        <div className="team-badge">
          <GraduationCap size={18} />
          <span>إعداد طالبات كلية التمريض • الفرقة الدراسية</span>
        </div>
        <p className="team-supervision">
          Faculty of Nursing for Girls • Al-Azhar University • Medical-Surgical Nursing Department
        </p>
      </div>

      <div className="team-grid" data-reveal="cards">
        {STUDENTS.map((student, idx) => (
          <div key={student.name} className="team-card">
            <div className="team-card-top">
              <div className="team-avatar">
                <span>{idx + 1}</span>
              </div>
              <div className="team-info">
                <h2 className="team-name">{student.name}</h2>
                <span className="team-role">{student.role}</span>
              </div>
            </div>
            <div className="team-card-footer">
              <span className="team-faculty-pill">Al-Azhar University</span>
              <Award size={16} className="team-award-icon" />
            </div>
          </div>
        ))}
      </div>

      <div className="team-bottom-banner" data-reveal="art">
        <Users size={20} className="team-banner-icon" />
        <p>
          Clinical seminar prepared under the academic curriculum of the Medical-Surgical Nursing Department,
          Faculty of Nursing for Girls, Al-Azhar University.
        </p>
      </div>
    </div>
  );
}
